import { RuuviInfo } from ".";

export interface RuuviTag {
  domain?: any;
  _events: Events;
  _eventsCount: number;
  _maxListeners: undefined;
  id: string;
  address: string;
  addressType: string;
  connectable: boolean;

  on: (eventName: string, cb: (data: RuuviInfo) => void) => void;
}

export interface Events {
}