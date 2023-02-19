import O365EventModel from "..//Models/IEventType";

export interface IMultiCalandarWpState {
  groupID: string;
  refreshing: boolean;
  loaded: boolean;
  events: O365EventModel [];
}
