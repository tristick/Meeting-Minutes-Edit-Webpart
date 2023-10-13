export interface IMeetingMinutesFormState { 
    title: string
    purposeofmeeting:string,
    managementsummary:string,
    mainminutes:string,
    actions:string,
    customer:string,
    meetingdate:Date,
    users:string [],
    attendeeDropdown:string,
    attendeesother:string,
    interestedPartiesexternal: string [],
    interestedPartiesexternalstr: string,
    allfieldsvalid:boolean,
    isSuccess: boolean,
    isfailure: boolean,
    meetingtitle:string,
    location:string,
    pmdocuments:string,
    msdocuments:string,
    mmdocuments:string,
    usersdisplayName:string []

}