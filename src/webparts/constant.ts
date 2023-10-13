import { IMeetingMinutesFormProps } from "./meetingMinuteEditWebpart/components/IMeetingMinutesFormProps";

export const LISTNAME ="Meeting Minutes";
export const METADATA_LISTNAME ="Metadata";
export const CONTACTS_LIST_NAME ="Contacts";
export const LIBRARYNAME = "MeetingMinutesSupportingFiles"//"Meeting Minutes Documents";
export const SUBMIT_REDIRECT = (props: IMeetingMinutesFormProps): URL => {return new URL(props.context.pageContext.web.absoluteUrl +"/SitePages/Home.aspx");};
export const CANCEL_REDIRECT = (props: IMeetingMinutesFormProps): URL => {return new URL(props.context.pageContext.web.absoluteUrl);};



export const modules = {  
    toolbar: [  
        [{  
            'header': [1, 2, 3, false]  
        }],  
        ['bold', 'italic', 'underline', 'strike', 'blockquote'],  
         
        [{  
            'list': 'ordered'  
        }, {  
            'list': 'bullet'  
        }, {  
            'indent': '-1'  
        }, {  
            'indent': '+1'  
        }],  
        ['image']  
        
    ],  
};
export const formats = ['header', 'bold', 'italic', 'underline', 'strike', 'blockquote', 'list', 'bullet', 'indent', 'image', 'background', 'color']; 