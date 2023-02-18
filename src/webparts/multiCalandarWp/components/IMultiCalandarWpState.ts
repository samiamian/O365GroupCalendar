import O365EventModel from "..//Models/IEventType";

export interface IMultiCalandarWpState {
  groupID: string;
  timeZone: string;
  events: O365EventModel [];
}
