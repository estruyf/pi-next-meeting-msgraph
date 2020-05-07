//@ts-check

import express from 'express';
import { Auth, AuthLogging } from './helpers';
import { MsGraphService } from './services';
import { formatRelative, format, addDays, subHours, parseJSON, differenceInMinutes } from 'date-fns';
import enGB from 'date-fns/locale/en-GB';
import ruuvi from 'node-ruuvitag';
import { RuuviTag, RuuviInfo, CalendarEvents, Availability, StatusInfo } from './models';
import fetch from 'node-fetch';

// Date formatting
const formatRelativeLocale = {
  lastWeek: "'Last' eeee 'at' HH:mm",
  yesterday: "'Yesterday at' HH:mm",
  today: "'Today at' HH:mm",
  tomorrow: "'Tomorrow at' HH:mm",
  nextWeek: "eeee 'at' HH:mm",
  other: 'HH:mm'
};

const locale = {
  ...enGB,
  formatRelative: (token: string) => formatRelativeLocale[token]
};

const app = express();

const MSGRAPH_URL = `https://graph.microsoft.com`;
const PORT = process.env.PORT || 1337;
const APP_ID = process.env.APPID || "";
const TODO_APP_ID = process.env.TODOAPPID || "";
const DEBUG = process.env.DEBUG ? process.env.DEBUG == "true" : false;
const STATUS_API = process.env.STATUSAPI || ""

const auth = new Auth(APP_ID, "Meeting");
const todoAuth = new Auth(TODO_APP_ID, "Todo");
let temperature = null;
let availability = null;
let todoTasks = null;
let nextMeeting = { title: "", time: "" };
let timeoutIdx: NodeJS.Timeout = null;
let authMsg: AuthLogging = {
  text: ""
};

app.use(express.json());
app.use(express.urlencoded({ extended: false }));

app.get('/get', (req, res) => res.send({ meeting: nextMeeting, availability, temperature, todoTasks }));
app.get('/auth', (req, res) => res.send(authMsg));
app.get('/status', async (req, res) => {
  const status = await getStatus();
  res.send({ status, availability });
});
app.get('/meeting', async (req, res) => {
  const meeting = await getMeeting();
  res.send({ meeting });
});
app.get('/todo', async (req, res) => {
  const tasks = await getTodo();
  res.send({ tasks });
});

app.get('/restart', (req, res) => {
  if (timeoutIdx) {
    clearTimeout(timeoutIdx);
    timeoutIdx = null;
  }
  startAuthentication();
  startTodoAuthentication();
  res.send("Authentication restarted");
});

app.listen(PORT, () => {
  console.log('Listening on port %s for inbound button push event notifications', PORT);
  startAuthentication();
  startTodoAuthentication();
});

/**
 * Starts the autentication flow
 */
const startAuthentication = () => {
  auth.ensureAccessToken(MSGRAPH_URL, authMsg, DEBUG).then(async (accessToken) => {
    if (accessToken) {
      console.log(`Calendar access token acquired.`);
      presencePolling();
    }
  });
}

/**
 * Starts todo autentication flow
 */
const startTodoAuthentication = () => {
  todoAuth.ensureAccessToken(MSGRAPH_URL, authMsg, DEBUG).then(async (accessToken) => {
    if (accessToken) {
      console.log(`Todo access token acquired.`);
      todoPolling();
    }
  });
}

/**
 * Keep calling the MS Graph to keep the token alive
 */
const presencePolling = async () => {
  await getMeeting();

  await getStatus();

  timeoutIdx = setTimeout(() => {
    presencePolling();
  }, 1 * 60 * 1000);
}

/**
 * Todo polling
 */
const todoPolling = async () => {
  await getTodo();

  timeoutIdx = setTimeout(() => {
    todoPolling();
  }, 1 * 60 * 1000);
}

/**
 * Retrieve the todo details
 */
const getTodo = async () => {
  try {
    const accessToken = await todoAuth.ensureAccessToken(MSGRAPH_URL, authMsg, DEBUG);
    let msGraphEndPoint = `${MSGRAPH_URL}/beta/me/outlook/tasks?$count=true&$select=subject&$filter=dueDateTime/datetime eq '${format(new Date(), "yyyy-MM-dd")}T00:00:00${encodeURIComponent(format(new Date(), "xxx"))}'`;

    if (DEBUG) {
      console.log(`Calling: ${msGraphEndPoint}`);
    }

    const todoItems: Tasks = await MsGraphService.get(`${msGraphEndPoint}`, accessToken, DEBUG);
    if (todoItems && todoItems.value && todoItems.value.length > 0) {
      console.log(`Todo`, todoItems);

      todoTasks = todoItems["@odata.count"];

      return todoItems;
    }

    todoTasks = null;
    return null;
  } catch (e) {
    console.error(e.message);
    todoTasks = null;
    return null;
  }
};

/**
 * Retrieve the meeting details
 */
const getMeeting = async () => {
  try {
    const accessToken = await auth.ensureAccessToken(MSGRAPH_URL, authMsg, DEBUG);
    if (accessToken) {
      let msGraphEndPoint = `${MSGRAPH_URL}/v1.0/me/calendarview?startdatetime=${format(subHours(new Date(), 1), "yyyy-MM-dd'T'HH:mm:ss")}%2B01:00&enddatetime=${format(addDays(new Date(), 1), "yyyy-MM-dd'T'HH:mm:ss")}%2B01:00&$select=subject,location,start&$top=5&$orderby=start/dateTime asc&$filter=isAllDay eq false`;

      if (DEBUG) {
        console.log(`Calling: ${msGraphEndPoint}`);
      }

      const calendarItems: CalendarEvents = await MsGraphService.get(`${msGraphEndPoint}`, accessToken, DEBUG);
      if (calendarItems && calendarItems.value && calendarItems.value.length > 0) {
        for (const event of calendarItems.value) {
          const eventDate = parseJSON(event.start.dateTime);
          const difference = differenceInMinutes(eventDate, new Date());
          
          if (DEBUG) {
            console.log(`Meeting "${event.subject}" time difference: ${difference}`);
          }

          if (difference > 0) {
            nextMeeting = {
              title: event.subject,
              time: formatRelative(parseJSON(event.start.dateTime), new Date(), { locale })
            };
            return calendarItems;
          }
        }        
      }
    }
    nextMeeting = {
      title: "",
      time: ""
    };
    return null;
  } catch (e) {
    console.error(e.message);
    nextMeeting = {
      title: "",
      time: ""
    };
    return null;
  }
};

/**
 * Get the status details
 */
const getStatus = async () => {
  try {
    if (STATUS_API) {
      const data = await fetch(STATUS_API);
      if (data && data.ok) {
        const status: StatusInfo = await data.json();
        if (DEBUG) {
          console.log(`Status:`, JSON.stringify(status));
        }

        if (status.red === 0 && status.green === 144 && status.blue === 0) {
          availability = Availability.Available;
        } else if (status.red === 255 && status.green === 191 && status.blue === 0) {
          availability = Availability.Away;
        } else if (status.red === 179 && status.green === 0 && status.blue === 0) {
          availability = Availability.Busy;
        } else {
          availability = Availability.Away;
        }

        return status;
      }
    }
    return null;
  } catch (e) {
    console.error(e.message);
    availability = Availability.Away;
    return null;
  }
};

/**
 * Ruuvi
 */
ruuvi.on('found', (tag: RuuviTag) => {
  if (DEBUG) {
    console.log(`Ruuvi tag:`, tag);
  }
  
  tag.on('updated', (data: RuuviInfo) => {
    if (data && data.temperature) {
      if (DEBUG) {
        console.log(`Ruuvi tag data:`, data);
      }
      temperature = data.temperature;
    }      
  });
});