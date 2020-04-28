export interface CalendarEvents {
  '@odata.context': string;
  '@odata.nextLink': string;
  value: Event[];
}

export interface Event {
  '@odata.etag': string;
  id: string;
  subject: string;
  start: Start;
  location: Location;
}

export interface Location {
  displayName: string;
  locationType: string;
  uniqueIdType: string;
  address: Address;
  coordinates: Address;
}

export interface Address {
}

export interface Start {
  dateTime: string;
  timeZone: string;
}