import O365EventModel from "..//Models/IEventType";

export interface IMultiCalandarWpState {
  groupID: string;
  refreshed: boolean;
  loading: boolean;
  events: O365EventModel [];
}
