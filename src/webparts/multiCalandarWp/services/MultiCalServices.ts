import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClient,SPHttpClient } from "@microsoft/sp-http";
import O365EventModel from "..//Models/IEventType";
import { Frequency, RRule, Weekday } from 'rrule';

export class MultiCalService {


   /** @param {Promise} promise
    * @returns {Promise} [ data, undefined ]
    * @returns {Promise} [ undefined, Error ]
    */
    private handle = <T>(promise: Promise<T>, defaultError: any = 'rejected'): Promise<T[] | [T, any]> => {
        return promise
          .then((data) => [data, undefined])
          .catch(error => ([undefined, error || defaultError]));
      }

    public async getGroupGUID(siteName: string, context: WebPartContext): Promise<string> {

        let groupGUID: string;
        const absoluteURL: string = context.pageContext.web.absoluteUrl;
        const SITE_ALL_PROPERTIES_URL = "/_api/web/allproperties";
        const endPointUrl: string = absoluteURL.substring(0, absoluteURL.lastIndexOf("/"))
            .concat("/")
            .concat(siteName)
            .concat(SITE_ALL_PROPERTIES_URL);

            let [response, responseErr] = await this.handle(context.spHttpClient.get(endPointUrl, SPHttpClient.configurations.v1));
            if(responseErr) throw new Error('Could not fetch Group ID details');

            let [jsonData, jsonDataErr] = await this.handle(response.json());
            if(jsonDataErr) throw new Error('Could not Group ID JSON Data');

            if (jsonData["GroupId"] !== undefined) {
                groupGUID = jsonData["GroupId"].toString();                
            }
            return groupGUID;
    }

    public async getCurrentSiteTimeZone(context: WebPartContext): Promise<string> {
        let timeZone: string;
        const absoluteURL: string = context.pageContext.web.absoluteUrl;
        const SITE_REGIONAL_SETTINGS_URL = "/_api/web/regionalsettings/timezone?$select=Description";
        const endPointUrl: string = absoluteURL.concat(SITE_REGIONAL_SETTINGS_URL);       
        
        let [data,dataerr] = await this.handle(context.spHttpClient.get(endPointUrl, SPHttpClient.configurations.v1));
        if(dataerr) throw new Error('Could not fetch timezone details');

        let [jsonData, jsonDataErr] = await this.handle(data.json());
        if(jsonDataErr) throw new Error('Could not get TimeZone JSON Data');

        if (jsonData !== undefined){
            timeZone = jsonData["Description"];
        }
        // if (timeZone.toLowerCase().indexOf('pacific time') > -1) {
        //     timeZone = 'America/Los_Angeles';
        //   }
        //   else if (timeZone.toLowerCase().indexOf('central time') > -1) {
        //     timeZone = 'America/Chicago';
        //   }
        //   else if (timeZone.toLowerCase().indexOf('eastern time') > -1) {
        //     timeZone = 'America/New_York';
        //   }
        //   else if (timeZone.toLowerCase().indexOf('mountain time') > -1) {
        //     timeZone = 'America/Denver';
        //   }
        //   else {
        //     timeZone = 'UTC';
        //   }
        return timeZone;
    }
    public async getAllGroupEvents(groupId: string,  context: WebPartContext, recursiveCall = false): Promise<O365EventModel[]> {

        let responseToReturn: O365EventModel[] = [];
        let responseToReturnErr: string;
        let options  = '?$select=subject,body,bodyPreview,organizer,attendees,start,end,location,recurrence';

        let [graphClient, graphClientErr] = await this.handle(context.msGraphClientFactory.getClient());
        if (graphClientErr) throw new Error("unable to get graphclient");

        let [tz,tzerror] = await this.handle(this.getCurrentSiteTimeZone(context));
        if (tzerror) throw new Error("unable to get timezone from regional settings");
        
       [responseToReturn, responseToReturnErr] = await this.handle(this.getRecursiveEvents(graphClient,groupId,responseToReturn,tz,options,false));
        if (responseToReturnErr) throw new Error(" No events found for groupID :: "+groupId);

        return responseToReturn;
    }

    private async getRecursiveEvents(client: MSGraphClient, groupId: string, groupEvents: any[], timeZone: string, options: string, isRecursiveCall: boolean = false): Promise <O365EventModel[]> {

        let eventURL: string = '';

        if (isRecursiveCall == false) {
            eventURL = `https://graph.microsoft.com/v1.0/groups/${groupId}/events${options}`;
        }
        else {
            eventURL = groupId;
        }       

        console.log(eventURL);

        await client.api(eventURL).get().then(async (data) => {

            if (data["@odata.nextLink"]) {

                await this.getRecursiveEvents(client, data["@odata.nextLink"], data, timeZone, options,  true);
            }
            
            data.value.map(respItem => {
            const attendents = {...respItem["attendees"]};

            if (respItem["recurrence"] == null){

                    groupEvents.push({
                        id: respItem["id"],
                        title: respItem["subject"],
                        bodyPreview: respItem["bodyPreview"],
                        start: new Date(respItem["start"]["dateTime"]), 
                        end: new Date(respItem["end"]["dateTime"]), 
                        timeZone: timeZone, 
                        organizerAddress: respItem["organizer"]["emailAddress"]["address"],
                        attendees: attendents,
                    });
                }
                else{
                groupEvents.push(...(this.getReoccuringEvents(respItem,timeZone,attendents)));
                }
            });
            
        }).catch(err => { console.log(err); });
        
        console.log(groupEvents);
        return groupEvents;
}

    private getReoccuringEvents(event: object, timeZone: string, attendents: any): O365EventModel[]{
        let reoccuringEvents: O365EventModel[] = [];

        // get the recurrence type
        let weekDays: string[] = [];
        let rruleWeekDays: Weekday[] = [];
        let frequency: Frequency;
        let recurrenceType: string = event["recurrence"]["pattern"]["type"];
        let rangeStart: string = event["recurrence"]["range"]["startDate"];
        let rangeEnd: string = event["recurrence"]["range"]["endDate"];
        let timeStart: string = event["start"]["dateTime"].split('T')[1];
        let timeEnd: string = event["end"]["dateTime"].split('T')[1];
        let rangeStartYear: number = parseInt(rangeStart.split('-')[0]);
        let rangeStartMonth: number = parseInt(rangeStart.split('-')[1]);
        let rangeStartDay: number = parseInt(rangeStart.split('-')[2]);
        let rangeEndYear: number = parseInt(rangeEnd.split('-')[0]);
        let rangeEndMonth: number = parseInt(rangeEnd.split('-')[1]);
        let rangeEndDay: number = parseInt(rangeEnd.split('-')[2]);
        let startHour: number = parseInt(timeStart.split(':')[0]);
        let startMin: number = parseInt(timeStart.split(':')[1]);
        let startSec: number = parseInt(timeStart.split(':')[2]);
        let endHour: number = parseInt(timeEnd.split(':')[0]);
        let endMin: number = parseInt(timeEnd.split(':')[1]);
        let endSec: number = parseInt(timeEnd.split(':')[2]);
        let secondsDiff = Math.abs(new Date(event["start"]["dateTime"]).getTime() - new Date(event["end"]["dateTime"]).getTime()) / 1000;


        ({ frequency, weekDays } = this.getFrequency(recurrenceType, frequency, event, weekDays));

        // if weekdays are selected then fetch the weekday names according to RRULE interface
        this.getWeekDayRules(weekDays, rruleWeekDays);

        const rule = new RRule({
            freq: frequency,
            dtstart: new Date(Date.UTC(rangeStartYear, rangeStartMonth-1, rangeStartDay, startHour, startMin, startSec, 0)),
            until: new Date(Date.UTC(rangeEndYear, rangeEndMonth-1, rangeEndDay, endHour, endMin, endSec, 0)),
            interval: event["recurrence"]["pattern"]["interval"],
            byweekday: rruleWeekDays
          });
          

        let allDates: Date[] = rule.all();

        if (allDates.length > 0) {
            allDates.forEach((date) => {
                let evtEndDateObj: Date = new Date();
                evtEndDateObj.setTime(date.getTime() + (secondsDiff * 1000));
                reoccuringEvents.push({
                    id: event["id"],
                    title: event["subject"],
                    bodyPreview: event["bodyPreview"],
                    start: date, 
                    end: evtEndDateObj,
                    timeZone: timeZone,
                    organizerAddress: event["organizer"]["emailAddress"]["address"],
                    attendees: attendents,
                });
            });
          }
        return reoccuringEvents;
    }

    private getWeekDayRules(weekDays: string[], rruleWeekDays: Weekday[]) {
        if (weekDays.length > 0) {
            weekDays.forEach((dayName) => {
                switch (dayName.toLocaleLowerCase()) {
                    case "monday":
                        rruleWeekDays.push(RRule.MO);
                        break;
                    case "tuesday":
                        rruleWeekDays.push(RRule.TU);
                        break;
                    case "wednesday":
                        rruleWeekDays.push(RRule.WE);
                        break;
                    case "thursday":
                        rruleWeekDays.push(RRule.TH);
                        break;
                    case "friday":
                        rruleWeekDays.push(RRule.FR);
                        break;
                    case "saturday":
                        rruleWeekDays.push(RRule.SA);
                        break;
                    case "sunday":
                        rruleWeekDays.push(RRule.SU);
                        break;
                }
            });
        }
    }

    private getFrequency(recurrenceType: string, frequency: Frequency, event: object, weekDays: string[]) {
        switch (recurrenceType) {
            case "daily":
                frequency = RRule.DAILY;
                break;
            case "weekly":
                frequency = RRule.WEEKLY;
                if (event["recurrence"]["pattern"]["daysOfWeek"]) {
                    weekDays = event["recurrence"]["pattern"]["daysOfWeek"];
                }
                break;
            case "monthly":
                frequency = RRule.MONTHLY;
                break;
            case "yearly":
                frequency = RRule.YEARLY;
                break;
            default:
                frequency = null;
                break;
        }
        return { frequency, weekDays };
    }


    //    // public async getRecursiveEvents(groupId: string, context: WebPartContext,): Promise <O365EventModel[]> {
    //         // const graph = graphfi().using(SPFx(context));
     
    //          let groupEvents: O365EventModel[] = [];
    //          let options  = '?$select=subject,body,bodyPreview,organizer,attendees,start,end,location,recurrence';
    //          let eventURL = `https://graph.microsoft.com/v1.0/groups/${groupId}/events${options}`;
             
             
    //          let [graphClient, graphClientErr] = await this.handle(context.msGraphClientFactory.getClient());
    //          if (graphClientErr) throw new Error("unable to get graphclient ");
     
    //          let [tz,tzerror] = await this.handle(this.getCurrentSiteTimeZone(context));
    //          if (tzerror) throw new Error("unable to get timezone from regional settings");
    //          // await graph.groups.getById(groupId).calendar.events().then(async (data) => {
    //          //     console.log(data);
    //          //     data.map(respItem => {
    //          //         const attendents: any[] = {...respItem["attendees"]};
         
    //          //         if (respItem["recurrence"] == null){
    //          //                 groupEvents.push({
    //          //                     id: respItem["id"],
    //          //                     title: respItem["subject"],
    //          //                     bodyPreview: respItem["bodyPreview"],
    //          //                     start: new Date(respItem["start"]["dateTime"]), 
    //          //                     end: new Date(respItem["end"]["dateTime"]), 
    //          //                     timeZone: tz, 
    //          //                     organizerAddress: respItem["organizer"]["emailAddress"]["address"],
    //          //                     attendees: attendents,
    //          //                 });
    //          //             }
    //          //             else{
    //          //                let a =  this.getReoccuringEvents(respItem,tz,attendents);
    //          //                groupEvents.push(...a);
    //          //             }
    //          //         });
    //          // }).catch(err => { console.log(err); });
     
    //          await graphClient.api(eventURL).get().then(async (data) => {
    //              data.value.map(respItem => {
    //              const attendents = {...respItem["attendees"]};
     
    //              if (respItem["recurrence"] == null){
    //                      groupEvents.push({
    //                          id: respItem["id"],
    //                          title: respItem["subject"],
    //                          bodyPreview: respItem["bodyPreview"],
    //                          start: new Date(respItem["start"]["dateTime"]), //.toLocaleString("en-US", {timeZone: "US/Mountain"})
    //                          end: new Date(respItem["end"]["dateTime"]), //.toLocaleString("en-US", {timeZone: "US/Mountain"})
    //                          timeZone: tz, //timeZone: respItem["end"]["timeZone"],
    //                          organizerAddress: respItem["organizer"]["emailAddress"]["address"],
    //                          attendees: attendents,
    //                      });
    //                  }
    //                  else{
    //                     let a =  this.getReoccuringEvents(respItem,tz,attendents);
    //                     groupEvents.push(...a);
    //                  }
    //              });
    //          }).catch(err => { console.log(err); });
             
    //          console.log(groupEvents);
    //          return groupEvents;
    //      }
     
}
const aService = new MultiCalService();
export default aService;