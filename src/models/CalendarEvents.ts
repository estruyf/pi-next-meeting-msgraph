interface CalendarEvents {
  '@odata.context': string;
  '@odata.nextLink': string;
  value: Event[];
}

interface Event {
  '@odata.etag': string;
  id: string;
  subject: string;
  start: Start;
  location: Location;
}

interface Location {
  displayName: string;
  locationType: string;
  uniqueIdType: string;
  address: Address;
  coordinates: Address;
}

interface Address {
}

interface Start {
  dateTime: string;
  timeZone: string;
}