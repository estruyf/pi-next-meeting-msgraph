# pi-next-meeting-msgraph

Connects to Microsoft Graph and fetches the next meeting

## Process

- Server
- Connect and keep the authentication context
- Fetch -> Returns the next meeting

## API to call

https://graph.microsoft.com/v1.0/me/calendarview?startdatetime=2020-04-25T10:23:17.728Z&enddatetime=2020-04-28T10:23:17.728Z&$select=subject,location,start&$top=1&$orderby=start/dateTime asc&$filter=isAllDay eq false


## Deamon

Install the following:

```
npm install pm2 -g
npm run build
pm2 start ./dist/app.js
```

Monitoring:

```
pm2 monit
```

Startup:

```
pm2 startup
```

Get the log output (to login):

```
pm2 logs
```

