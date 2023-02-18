export default interface O365EventModel {
    id: string;
    title: string;
    bodyPreview: string;
    organizerAddress: string;
    attendees: Array<{type: string; name: string; email: string;}>;
    start:any;
    timeZone: string;
    end: any;
  }