//@ts-check

import express from 'express';
import { Auth } from './helpers';
import { MsGraphService } from './services';
import { formatRelative, format, addDays, parseJSON } from 'date-fns';

const app = express();

const MSGRAPH_URL = `https://graph.microsoft.com`;
const PORT = process.env.PORT || 1337;
const APP_ID = process.env.APPID || "";
const DEBUG = process.env.DEBUG ? process.env.DEBUG == "true" : false;

const auth = new Auth(APP_ID);
let nextMeeting = { title: "", time: "" };
let timeoutIdx: NodeJS.Timeout = null;

app.use(express.json());
app.use(express.urlencoded({ extended: false }));

app.get('/get', (req, res) => res.send(nextMeeting));
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
  auth.ensureAccessToken(MSGRAPH_URL, DEBUG).then(async (accessToken) => {
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
  const accessToken = await auth.ensureAccessToken(MSGRAPH_URL, DEBUG);
  if (accessToken) {
    const msGraphEndPoint = `v1.0/me/calendarview?startdatetime=${format(new Date(), "yyyy-MM-dd'T'HH:mm:ss")}&enddatetime=${format(addDays(new Date(), 1), "yyyy-MM-dd'T'HH:mm:ss")}&$select=subject,location,start&$top=1&$orderby=start/dateTime asc&$filter=isAllDay eq false`;
    const calendarItems: CalendarEvents = await MsGraphService.get(`${MSGRAPH_URL}/${msGraphEndPoint}`, accessToken, DEBUG);
    if (calendarItems && calendarItems.value && calendarItems.value.length > 0) {
      const event = calendarItems.value[0];
      nextMeeting = {
        title: event.subject,
        time: formatRelative(parseJSON(event.start.dateTime), new Date())
      };
    } else {
      nextMeeting = {
        title: "",
        time: ""
      };
    }
  }

  timeoutIdx = setTimeout(() => {
    presencePolling();
  }, 1 * 60 * 1000);
}