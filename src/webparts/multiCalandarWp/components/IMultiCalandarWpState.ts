import O365EventModel from "..//Models/IEventType";

export interface IMultiCalandarWpState {
  groupID: string;
  isGroupIDValid: boolean;
  refreshed: boolean;
  dataLoading: boolean;
  calendarLoading: boolean;
  events: O365EventModel [];
}
