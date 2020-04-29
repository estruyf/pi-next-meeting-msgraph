//@ts-check

import express from 'express';
import { Auth, AuthLogging } from './helpers';
import { MsGraphService } from './services';
import { formatRelative, format, addDays, subHours, parseJSON } from 'date-fns';
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
const DEBUG = process.env.DEBUG ? process.env.DEBUG == "true" : false;
const STATUS_API = process.env.STATUSAPI || ""

const auth = new Auth(APP_ID);
let temperature = null;
let availability = null;
let nextMeeting = { title: "", time: "" };
let timeoutIdx: NodeJS.Timeout = null;
let authMsg: AuthLogging = {
  text: ""
};

app.use(express.json());
app.use(express.urlencoded({ extended: false }));

app.get('/get', (req, res) => res.send({ meeting: nextMeeting, availability, temperature }));
app.get('/auth', (req, res) => res.send(authMsg));

app.get('/restart', (req, res) => {
  if (timeoutIdx) {
    clearTimeout(timeoutIdx);
    timeoutIdx = null;
  }
  startAuthentication();
  res.send("Authentication restarted");
});

app.listen(PORT, () => {
  console.log('Listening on port %s for inbound button push event notifications', PORT);
  startAuthentication();
});

/**
 * Starts the autentication flow
 */
const startAuthentication = () => {
  auth.ensureAccessToken(MSGRAPH_URL, authMsg, DEBUG).then(async (accessToken) => {
    if (accessToken) {
      console.log(`Access token acquired.`);
      presencePolling();
    }
  });
}

/**
 * Keep calling the MS Graph to keep the token alive
 */
const presencePolling = async () => {
  const accessToken = await auth.ensureAccessToken(MSGRAPH_URL, authMsg, DEBUG);
  if (accessToken) {
    const msGraphEndPoint = `v1.0/me/calendarview?startdatetime=${format(subHours(new Date(), 1), "yyyy-MM-dd'T'HH:mm:ss")}%2B01:00&enddatetime=${format(addDays(new Date(), 1), "yyyy-MM-dd'T'HH:mm:ss")}%2B01:00&$select=subject,location,start&$top=1&$orderby=start/dateTime asc&$filter=isAllDay eq false`;

    if (DEBUG) {
      console.log(`Calling: ${msGraphEndPoint}`);
    }

    const calendarItems: CalendarEvents = await MsGraphService.get(`${MSGRAPH_URL}/${msGraphEndPoint}`, accessToken, DEBUG);
    if (calendarItems && calendarItems.value && calendarItems.value.length > 0) {
      const event = calendarItems.value[0];
      nextMeeting = {
        title: event.subject,
        time: formatRelative(parseJSON(event.start.dateTime), new Date(), { locale })
      };
    } else {
      nextMeeting = {
        title: "",
        time: ""
      };
    }
  }

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
    }
  }

  timeoutIdx = setTimeout(() => {
    presencePolling();
  }, 1 * 60 * 1000);
}

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