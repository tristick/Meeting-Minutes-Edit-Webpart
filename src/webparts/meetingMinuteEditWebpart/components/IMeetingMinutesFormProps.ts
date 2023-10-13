import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMeetingMinutesFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  siteUrl:string
 
}
export interface ICustomer {

  Title:string;

}

export interface IListMM {

  Title: string,
        MeetingTitle:string,
        Customer: string,
        Location: string,
        MeetingDate: Date,
        AttendeesMOLEAId: any,
        AttendeesCustomer: string,
        AttendeesOther: string,
        PurposeofMeetingDocuments: string,
        ManagementSummaryDocuments:string,
        MainMinutesDocuments:string
        AttendeesMOLEADisplayNames:string
 

  
}
