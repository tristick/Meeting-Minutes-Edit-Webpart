import * as React from 'react';
import "@pnp/sp/folders";
import styles from './MeetingMinutesForm.module.scss';
import { IListMM, IMeetingMinutesFormProps } from './IMeetingMinutesFormProps';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as formconst from "../../constant";
import * as ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css'; 
import { IMeetingMinutesFormState } from './IMeetingMinutesFormState';
//,Label
import { IStyleFunctionOrObject, ITextFieldStyleProps, ITextFieldStyles, MessageBar, MessageBarType, PrimaryButton, SpinnerSize, Stack, TextField } from 'office-ui-fabric-react';
import { DateConvention, DateTimePicker, ListItemPicker } from '@pnp/spfx-controls-react';
import { getCustomerItem, getItem, readFile, submitDataAndGetId, updateData, uploadAttachment } from '../../../services/formservices';
import ReactDOM from 'react-dom';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../../pnpjsconfig';
import "@pnp/sp/files";
import { IHttpClientOptions, HttpClient } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react';

const textFieldStyles: Partial<ITextFieldStyles> = {
  field: {
    width: '600px', // Adjust the desired width
  },
};

let isemailInvalid:boolean = false;
let streamerror:boolean =false;
let isbuttondisbled : boolean = false;
let buttontext : string = "Submit";
let isselectedattendees:boolean = false ;
let listId: number;
let customerreference:string;
let isLoading : boolean = false;
let pmcontents:string = "";
let mmcontents:string ="";
let actcontents:string = "";
let mscontents:string=""

export default class MeetingMinuteEditWebpart extends React.Component<IMeetingMinutesFormProps, IMeetingMinutesFormState> {
  private pmdt: DataTransfer; 
  private msdt: DataTransfer; 
  private mmdt: DataTransfer; 


  filesNamesRef: React.RefObject<HTMLSpanElement>;

  constructor(props: IMeetingMinutesFormProps, state: IMeetingMinutesFormState) {  
    super(props); 
    this.pmdt = new DataTransfer();
    this.msdt = new DataTransfer();
    this.mmdt = new DataTransfer();
    this.filesNamesRef = React.createRef(); 
    this.state = {  
      title: "",
      purposeofmeeting:"",
      managementsummary:"",
      mainminutes:"",
      actions:"",
      customer:"",
      meetingdate:new Date(),
      users:[],
      attendeeDropdown:"",
      attendeesother:"",
      interestedPartiesexternal: [],
      interestedPartiesexternalstr:"",
      allfieldsvalid:true,
      isSuccess: false,
      isfailure:false,
      meetingtitle:"",
      location:"",
      pmdocuments:"",
      msdocuments:"",
      mmdocuments:"",
      usersdisplayName:[]

    }
  
  }

  public componentDidMount = async ()=>
  {
   
    this.fetchCustomer();
    //console.log(this.props.context.pageContext.site);
    const url = document.location.search //window.location.href
    const urlParams = new URLSearchParams (url);
    const paramtitle = urlParams.get("Title");
    console.log(paramtitle);
    let filepath = this.props.siteUrl+ "/" + formconst.LIBRARYNAME+ "/" + paramtitle
  //const filepath = formconst.LIBRARYNAME+ "/" + this.state.searchkey
    filepath = filepath.replace(/^.*?\/sites/,'/sites')
   pmcontents = await readFile(this.props,filepath +"/"+"PurposeOfMeeting.txt")
    this.setState({purposeofmeeting:decodeURIComponent(pmcontents)})

    mscontents = await readFile(this.props,filepath +"/"+"ManagementSummary.txt")
    this.setState({managementsummary:decodeURIComponent(mscontents)})
    
    //console.log(this.state.purposeofmeeting)
    mmcontents= await readFile(this.props,filepath +"/"+"MainMinutes.txt")
    this.setState({mainminutes:decodeURIComponent(mmcontents)})

  actcontents= await readFile(this.props,filepath +"/"+"Actions.txt")
    this.setState({actions:decodeURIComponent(actcontents)})

    const item: IListMM[] = await getItem(this.props,paramtitle);
    console.log('Fetched item:', item);
    const parsedDate = new Date((item[0].MeetingDate).toString());
    
 
    
    this.setState({

    title: item[0].Title,
    meetingtitle:item[0].MeetingTitle,
    //customer:item[0].Customer,
    location:item[0].Location,
    meetingdate:parsedDate,
    users:(item[0].AttendeesMOLEADisplayNames).split(','),
    attendeeDropdown:item[0].AttendeesCustomer,
    attendeesother:item[0].AttendeesOther,
    //pmdocuments:item[0].PurposeofMeetingDocuments,
    //msdocuments:item[0].ManagementSummaryDocuments,
    //mmdocuments:item[0].MainMinutesDocuments, 
    //usersdisplayName:item[0].AttendeesMOLEADisplayNames
  })
   
  }

  fetchCustomer = async () => {
    
      const customerItem:any = await getCustomerItem(this.props);
      console.log(customerItem)
      try {
        const customer = customerItem[0].Title
        console.log(customer)
        const customerRef = "MM";//customerItem[0].RefCode;
        console.log(customerRef)
      this.setState({customer:customer});
      customerreference = customerRef

   
    } catch (error) {
      console.error('Error fetching customer items:', error);
    }
  };

  private onpurposeofmeetingchange = (newText: string) => {
    
    this.setState({purposeofmeeting:newText});
   
    return newText;
  }
  private onmanagementsummarychange = (newText: string) => {
    this.setState({managementsummary:newText});
   
    return newText;
  }
  private onmainminuteschange = (newText: string) => {
    this.setState({mainminutes:newText});
   
    return newText;
  }
  private onactionschange = (newText: string) => {
    this.setState({actions:newText});
   
    return newText;
  }
  private _onchangedmeetingDate=(mdate: any): void =>{  
    this.setState({ meetingdate: mdate }); 

  }

  public _getPeoplePickerItems=(items: any[]) =>{  
    console.log(items)

    if(items.length>0){
      let selectedUsers: string[] = [];
      let selectedUsersdispName: string[] = [];
       items.map((item) => {
         selectedUsers.push(item.id);
         selectedUsersdispName.push(item.text)
       });
        this.setState({users: selectedUsers});
        this.setState({usersdisplayName:selectedUsersdispName})
       console.log('users:',selectedUsers)  
      isselectedattendees  = true;
      //console.log('Items new:', userid );
    }else{
      
      isselectedattendees  = false;

    }
     
       
       
  }
  
  /*private _onattendesSelectedItem=(data: { key: string; name: string }[])=> {
    
    if(data.length == 0 ){
      this.setState({attendeeDropdown:""})
    }else{
    this.setState({attendeeDropdown:data[0].name as string})
   
    }
  }*/

  private _onattendesSelectedItem=(data: { key: string; name: string }[])=> {

    console.log(data)
    
    if(data.length == 0 ){
      this.setState({attendeeDropdown:""})
    }else{
      let selectedUsers: string[] = [];
       data.map((item) => {
         selectedUsers.push(item.name);
        
       }); 
    this.setState({attendeeDropdown:(JSON.stringify(selectedUsers)).slice(1, -1).replace(/"/g, '').replace(/,/g, ", ")})
    
    console.log('attendeeusers:',selectedUsers)  
   
    }
  }

  private _onmeetingtitle=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
    //isemptybaf=isEmpty(newText)
    this.setState({meetingtitle:newText})
  
  }
  private _onlocation=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
    //isemptybaf=isEmpty(newText)
    this.setState({location:newText})
  
  }

  private onchangeattendeesother=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
    this.setState({attendeesother:newText})
  }

  private handleAddattendee = () => {
   const { attendeesother, interestedPartiesexternal } = this.state;
    if (attendeesother.trim() !== ''&& /^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/.test(attendeesother)) {
      const updatedParties = [...interestedPartiesexternal, attendeesother]
      console.log(updatedParties)
  
      this.setState({ interestedPartiesexternal: updatedParties, attendeesother: '', interestedPartiesexternalstr:(JSON.stringify(updatedParties)).slice(1, -1).replace(/"/g, '').replace(/,/g, ", ")});
      isemailInvalid = false;
    } else{
  
      isemailInvalid = true;
      this.setState({attendeesother:"",allfieldsvalid:false})
  
    }
  }

  private purposeofmeetinghandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
   
  const filesNames = document.querySelector<HTMLSpanElement>('#purposeofmeetingfilesList > #purposeofmeetingfiles-names');
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
          <span className="name">{e.target.files.item(i).name}</span>
  <span className="file-delete">
     <button> Remove</button>
  </span>
  <br/>
        </span>
      );
  
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
     
      }
    }
  
    for (let file of e.target.files as any) {
      this.pmdt.items.add(file);
    }
  
    e.target.files = this.pmdt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.pmdt.items.length; i++) {
          if (name === this.pmdt.items[i].getAsFile()?.name) {
            this.pmdt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.pmdt.files;
  
      });
    });
  };

  private managementsummaryhandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
   
    const filesNames = document.querySelector<HTMLSpanElement>('#managementsummaryfilesList > #managementsummaryfiles-names');
      for (let i = 0; i < e.target.files.length; i++) {
        let fileBloc = (
          <span key={i} className="file-block">
            <span className="name">{e.target.files.item(i).name}</span>
    <span className="file-delete">
       <button> Remove</button>
    </span>
    <br/>
          </span>
        );
    
        if (filesNames) {
          const fileBlocContainer = document.createElement('div');
          ReactDOM.render(fileBloc, fileBlocContainer);
          filesNames?.appendChild(fileBlocContainer.firstChild);
       
        }
      }
    
      for (let file of e.target.files as any) {
        this.msdt.items.add(file);
      }
    
      e.target.files = this.msdt.files;
    
      document.querySelectorAll('span.file-delete').forEach((button) => {
        button.addEventListener('click', () => {
          let name = button.nextSibling.textContent;
    
          (button.parentNode as HTMLElement)?.remove();
    
          for (let i = 0; i < this.msdt.items.length; i++) {
            if (name === this.msdt.items[i].getAsFile()?.name) {
              this.msdt.items.remove(i);
              continue;
            }
          }
    
          e.target.files = this.msdt.files;
    
        });
      });
    };

    private mainminuteshandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
   
      const filesNames = document.querySelector<HTMLSpanElement>('#mainminutesfilesList > #mainminutesfiles-names');
        for (let i = 0; i < e.target.files.length; i++) {
          let fileBloc = (
            <span key={i} className="file-block">
              <span className="name">{e.target.files.item(i).name}</span>
      <span className="file-delete">
         <button> Remove</button>
      </span>
      <br/>
            </span>
          );
      
          if (filesNames) {
            const fileBlocContainer = document.createElement('div');
            ReactDOM.render(fileBloc, fileBlocContainer);
            filesNames?.appendChild(fileBlocContainer.firstChild);
         
          }
        }
      
        for (let file of e.target.files as any) {
          this.mmdt.items.add(file);
        }
      
        e.target.files = this.mmdt.files;
      
        document.querySelectorAll('span.file-delete').forEach((button) => {
          button.addEventListener('click', () => {
            let name = button.nextSibling.textContent;
      
            (button.parentNode as HTMLElement)?.remove();
      
            for (let i = 0; i < this.mmdt.items.length; i++) {
              if (name === this.mmdt.items[i].getAsFile()?.name) {
                this.mmdt.items.remove(i);
                continue;
              }
            }
      
            e.target.files = this.mmdt.files;
      
          });
        });
      };

      private _createItem  =async (props:IMeetingMinutesFormProps):Promise<void>=>{
        //1048576 bytes = 1MB
        //|| (this.state.attendeeDropdown).length == 0
        //let allRichtextsizebinary = new Blob([this.state.purposeofmeeting]).size + new Blob([this.state.managementsummary]).size + new Blob([this.state.mainminutes]).size + new Blob([this.state.actions]).size;
        //|| new Blob([this.state.purposeofmeeting]).size >1000000|| new Blob([this.state.managementsummary]).size>1000000 ||
     //new Blob([this.state.mainminutes]).size>1000000 || new Blob([this.state.actions]).size>1000000 || allRichtextsizebinary>1048000
      if(isselectedattendees==false || isEmpty(this.state.meetingtitle)||isEmpty(this.state.location) || isEmpty(this.state.purposeofmeeting) || 
      isEmpty(this.state.mainminutes) 
      )
          {
          this.setState({allfieldsvalid:false}) ; 
          console.log(this.state.allfieldsvalid)
          
          return;
          }else {
            this.setState({allfieldsvalid:true}) ; 
            isbuttondisbled = true;
            buttontext = "Saving...";
            isLoading = true;
          }

          let folderUrl: string;

          const data = {
            Title: 'New Item creation in process',
            //PurposeOfMeeting: this.state.purposeofmeeting,
           // ManagementSummary: this.state.managementsummary,
            //MainMinutes: this.state.mainminutes,
            //Actions: this.state.actions,
           
          
         }; 

         submitDataAndGetId(this.props,data).then(async (itemId: any) => {
          listId = itemId   
          console.log(`Item created with ID: ${itemId}`);
  
          //Request ID format
          let now = new Date();
          let options: Intl.DateTimeFormatOptions = {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        };
        let listIdstr
         if(listId < 1000 && listId > 99){
          listIdstr = "0"+String(listId)
        }else if(listId < 100 && listId > 9){
          listIdstr ="00"+String(listId)
        } else if(listId < 10) {
          listIdstr ="000"+String(listId)
        }else{
          listIdstr = String(listId)
        }

       /* if(listId < 1000 && listId > 100){
          listIdstr = "0"+String(listId)
        }else if(listId < 100){
          listIdstr ="00"+String(listId)
        } else if(listId < 10) {
          listIdstr ="000"+String(listId)
        }else{
          listIdstr = String(listId)
        }*/
        
        console.log(listIdstr)
        let formattedDate = now.toLocaleDateString('en-GB', options).replace(/\//g, '');;
        let lastitemid = (listIdstr)+"-"+customerreference +"-" +formattedDate.toString();
  
     folderUrl =  formconst.LIBRARYNAME +"/" + lastitemid
      this.setState({title:lastitemid})
      
          
     
    }).then(async () => {
      
      await upload();
      await writeFile();
      // Update the item
      const updatedData = {
        Title: this.state.title,
        MeetingTitle:this.state.meetingtitle,
        Customer: this.state.customer,
        Location: this.state.location,
        MeetingDate: this.state.meetingdate,
        AttendeesMOLEAId: this.state.users,
        AttendeesCustomer: this.state.attendeeDropdown,
        AttendeesOther: this.state.interestedPartiesexternalstr,
        PurposeofMeetingDocuments: this.state.pmdocuments,
        ManagementSummaryDocuments:this.state.msdocuments,
        MainMinutesDocuments:this.state.mmdocuments,
        AttendeesMOLEADisplayNames:(JSON.stringify(this.state.usersdisplayName)).slice(1, -1).replace(/"/g, '').replace(/,/g, ", "),
      };
      return updateData(this.props,listId, updatedData);
    })
     .then(async () => {
      //console.log('Item Updated successfully');
      // Perform any further actions if needed
      await callflow()
      /* isbuttondisbled = false;
      buttontext = "Submit"
      this.setState({ isSuccess: true });
    
    window.open(formconst.SUBMIT_REDIRECT(props),"_self") */
    }) 
    .catch((error: any) => {
      
      /*var obj = JSON.stringify(error);
    
      if(obj.indexOf("400") !== -1)
      {    console.log("mATCH FOUND")
            streamerror = true;
           this.setState({allfieldsvalid:false}) 
    }
  
      else{*/
      isbuttondisbled = false;
      buttontext = "Submit";
      isLoading = false;
      this.setState({ isfailure: true });
      console.log('Error:', error);//}
    });
    const callflow = () =>{  
      //const postURL = "https://prod-30.centralindia.logic.azure.com:443/workflows/9faa10085fbd45898a35ad4a0ef35c96/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=HJoQXT0cEgKWnMy2MJOjM3lU4HTNOpK_nJaKw-Zo7Og";  
      //const postURL = "https://prod-04.uksouth.logic.azure.com:443/workflows/ee066a883d5d4963b219ed3f0d019eb3/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=x6UwhOGMRsGdJfX5zNyYh6N0tII5UJs4dkd4qWsWyMQ";//SKP
      //const postURL = "https://prod-21.westeurope.logic.azure.com:443/workflows/8b7d8eaf80f448d4ab98b79d3dffcba5/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=zJuPQYLWgvrdVrwelM3keIZxow7g60I8tYdgrSSNotw";//KWel
      const postURL = "https://prod-22.uksouth.logic.azure.com:443/workflows/59d54402715d43dfa56735dbc42b03f9/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=OzFad8oCWSKEq56j7IVpq4AkQmZZS3mMkeoPURLdUBA";//Client
      const body: string = JSON.stringify({  
        //'emailaddress': "kastuv@k6931.onmicrosoft.com",  
        'meetingID': this.state.title,  
        'customer': this.state.customer,  
      });
      
     
      const requestHeaders: Headers = new Headers();  
      
      requestHeaders.append('Content-type', 'application/json');  
     
      const httpClientOptions: IHttpClientOptions = {  
        body: body,  
        headers: requestHeaders  
      };  
      
     
      
     props.context.httpClient.post(postURL,HttpClient.configurations.v1,httpClientOptions)  
    
        .then(() => {  
          isbuttondisbled = false;
          buttontext = "Submit";
          isLoading = false;
          this.setState({ isSuccess: true });
        
        window.open(formconst.SUBMIT_REDIRECT(props),"_self")
         
        });  
      }
  
   const writeFile = async () => {
 
    const purposefileContent = encodeURIComponent(this.state.purposeofmeeting);
    const purposefileName = "PurposeOfMeeting.txt";
   
    try {
      await uploadAttachment(this.props,folderUrl,purposefileName, purposefileContent,this.state.title)
     
   } catch (err) {
     console.error("Error uploading file:", err);
   }
   const managementSummaryfileContent = encodeURIComponent(this.state.managementsummary);
    const managementSummaryfileName = "ManagementSummary.txt";
    
    try {
      await uploadAttachment(this.props,folderUrl,managementSummaryfileName, managementSummaryfileContent,this.state.title)
     
   } catch (err) {
     console.error("Error uploading file:", err);
   }
   const mainminutesfileContent = encodeURIComponent(this.state.mainminutes);
    const mainminutesfileName = "MainMinutes.txt";
    
   
    try {
      await uploadAttachment(this.props,folderUrl,mainminutesfileName, mainminutesfileContent,this.state.title)
     
   } catch (err) {
     console.error("Error uploading file:", err);
   }
   const actionsfileContent = encodeURIComponent(this.state.actions);
    const actionsfileName = " Actions.txt";
    
    try {
      await uploadAttachment(this.props,folderUrl,actionsfileName, actionsfileContent,this.state.title)
     
   } catch (err) {
     console.error("Error uploading file:", err);
   }
  }
   
  const upload = async () => {
  
      console.log(folderUrl)
      const _sp :SPFI = getSP(props.context) ;
      let strbgurl = "";
      let vstrbgurl = "";
      let ostrbgurl = "";
      _sp.web.folders.addUsingPath(folderUrl);
      // bgfiles
      
      let bgfileurl = [];

      let bginput = document.getElementById("purposeofmeetingattachment") as HTMLInputElement;
  
      console.log(bginput.files);
    
      if (bginput.files.length > 0) {
        let bgfiles = bginput.files;
      
        for (var i = 0; i < bgfiles.length; i++) {
          let bgfile = bginput.files[i];
          console.log("bgfile",bgfile)
          bgfileurl.push(this.props.siteUrl+ "/" + folderUrl + "/" +bgfile.name);
          //console.log()
          try {
            await uploadAttachment(this.props,folderUrl,bgfile.name, bgfile,this.state.title)
           
          } catch (err) {
            console.error("Error uploading file:", err);
          }
        }
        let convertedStr = bgfileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
         strbgurl = convertedStr.toString().replace(/,/g, ", ");
          //console.log(strbgurl);
          this.setState({ pmdocuments: strbgurl });
      }
        
       else {
        console.log("No file selected for upload.");
      }
      // vfiles
      let vfileurl = [];
      let vinput = document.getElementById("managementsummaryattachment") as HTMLInputElement;
 
      console.log(vinput.files);
      if (vinput.files.length > 0) {
        let vfiles = vinput.files;
      
        for (var i = 0; i < vfiles.length; i++) {
          let vfile = vinput.files[i];
          console.log("vfile",vfile)
          vfileurl.push(this.props.siteUrl + "/" + folderUrl + "/" + vfile.name);
          try {
            await uploadAttachment(this.props,folderUrl,vfile.name, vfile,this.state.title)
           
         } catch (err) {
           console.error("Error uploading file:", err);
         }
        }
        let vconvertedStr = vfileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
       vstrbgurl = vconvertedStr.toString().replace(/,/g, ", ");
      //console.log(vstrbgurl);
      this.setState({ msdocuments: vstrbgurl });
      
      } else {
        console.log("No file selected for upload.");
        
      }
      
    
      // ofiles
      let ofileurl = [];
      let oinput = document.getElementById("mainminutesattachment") as HTMLInputElement;
  
      console.log(oinput.files);
     
      if (oinput.files.length > 0) {
        let ofiles = oinput.files;
     
        for (var i = 0; i < ofiles.length; i++) {
          let ofile = oinput.files[i];
          console.log("ofile",ofile)
          ofileurl.push(this.props.siteUrl+ "/" + folderUrl + "/" + ofile.name);
          try {
            await uploadAttachment(this.props,folderUrl,ofile.name, ofile,this.state.title)
           
         } catch (err) {
           console.error("Error uploading file:", err);
         }
        }
        let oconvertedStr = ofileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
         ostrbgurl = oconvertedStr.toString().replace(/,/g, ", ");
        //console.log(ostrbgurl);
        this.setState({ mmdocuments: ostrbgurl });
        
      } else {
        console.log("No file selected for upload.");
        
      }
  
      
     
    }
      
      

    }


      private cancel =()=>{
        window.open(formconst.CANCEL_REDIRECT(this.props),"_self");
      }

      private _resetrichtext = () =>{
 
        this.setState({purposeofmeeting:"", managementsummary:"",mainminutes:"", actions:"",allfieldsvalid:true})
        streamerror = false;
        isbuttondisbled = false;
        buttontext = "Submit"
      
      }


  public render(): React.ReactElement<IMeetingMinutesFormProps> {
    const {interestedPartiesexternal } = this.state;
    let EmailFieldErrorMessage: JSX.Element | null
    let imageFieldErrorMessage: JSX.Element | null
    let successMessage : JSX.Element | null
    let meetingtitleFieldErrorMessage : JSX.Element | null
    let locationFieldErrorMessage : JSX.Element | null
    let pmFieldErrorMessage : JSX.Element | null
    let mmFieldErrorMessage : JSX.Element | null
    let attendeeFieldErrorMessage : JSX.Element | null
    // let attcustFieldErrorMessage : JSX.Element | null
    let FormFieldErrorMessage : JSX.Element | null
    let allRichtextsizebinaryErrorMessage: JSX.Element | null
    let failureMessage : JSX.Element | null
    
    if(!this.state.allfieldsvalid){
      
      attendeeFieldErrorMessage = (isselectedattendees==false) ?
        <MessageBar messageBarType={MessageBarType.error}>Attendees (MOLEA)
        is required.</MessageBar>
        : null;
      meetingtitleFieldErrorMessage = isEmpty(this.state.meetingtitle) ?
        <MessageBar messageBarType={MessageBarType.error}>Meeting Title is required.</MessageBar>
        : null;
      locationFieldErrorMessage = isEmpty(this.state.location) ?
        <MessageBar messageBarType={MessageBarType.error}>Location is required.</MessageBar>
        : null;  
        pmFieldErrorMessage = isEmpty(this.state.purposeofmeeting) ?
        <MessageBar messageBarType={MessageBarType.error}>Purpose of Meeting
        is required.</MessageBar>
        : null; 
        mmFieldErrorMessage = isEmpty(this.state.mainminutes) ?
        <MessageBar messageBarType={MessageBarType.error}>Main Minutes
        is required.</MessageBar>
        : null; 
      
/*       attcustFieldErrorMessage = (this.state.attendeeDropdown).length == 0  ?
        <MessageBar messageBarType={MessageBarType.error}>Attendees (Customer)
        is required.</MessageBar>
        : null; */


      EmailFieldErrorMessage= isemailInvalid ?
      <MessageBar messageBarType={MessageBarType.error}>Please enter a valid email address.</MessageBar>
      : null;

      imageFieldErrorMessage = streamerror ? <MessageBar messageBarType={MessageBarType.blocked} isMultiline={false} onDismiss={this._resetrichtext} dismissButtonAriaLabel="Close"
      truncated={true} overflowButtonAriaLabel="See more">Stream size exceeds the allowed limit. Note that the image size in the rich text field should be less than 80 KB .
      On closing the dialog will reset the rich text field values </MessageBar>: null;

        //1048576 bytes = 1MB
        //let allRichtextsizebinary = new Blob([this.state.purposeofmeeting]).size + new Blob([this.state.managementsummary]).size + new Blob([this.state.mainminutes]).size + new Blob([this.state.actions]).size;
        //allRichtextsizebinaryErrorMessage = allRichtextsizebinary>1048000?
        //<MessageBar messageBarType={MessageBarType.error}>New request size exceeds allowed limit. Recommend reducing image sizes. </MessageBar>
        //: null;

       FormFieldErrorMessage= 
       <MessageBar messageBarType={MessageBarType.error}>Please provide all required information and submit the form.</MessageBar>
    
    }

    successMessage = this.state.isSuccess ?
    <MessageBar messageBarType={MessageBarType.success}>Meeting Id : {this.state.title} submitted successfully.</MessageBar>
    : null;

    failureMessage = this.state.isfailure ?
    <MessageBar messageBarType={MessageBarType.error}>Error : Contact your IT</MessageBar>
    : null;
 
   return (
      <section>
        <div>
{/*         <div>
          <p><span className={styles.required}><b>*</b></span> Required.</p>
          </div> */}
          <p className={styles.heading}>Overview</p>
        {/* <p className={styles.formlabel}>Customer<span className={styles.required}> *</span></p> */}
        <p className={styles.formlabel}>Customer</p>
        <p>{this.state.customer}</p>
        {/* <Label>{this.state.customer}</Label> */}

        <p className={styles.formlabel}>Meeting Title<span className={styles.required}> *</span></p>  
        <TextField value={this.state.meetingtitle} onChange={this._onmeetingtitle} />{meetingtitleFieldErrorMessage}

        <table>
          <tr ><td className={styles.tabltr}>
        <p className={styles.formlabel}>Meeting Date<span className={styles.required}> *</span></p> 
        <DateTimePicker 
          dateConvention={DateConvention.Date}
          value={this.state.meetingdate}  
          onChange={this._onchangedmeetingDate}
          allowTextInput = {false}  
          showLabels = {false}/>
          </td>
          <td width={'600px;'}></td>
          </tr>
      </table>

        <p className={styles.formlabel}>Location<span className={styles.required}> *</span></p>  
        <TextField value={this.state.location} onChange={this._onlocation}/>{locationFieldErrorMessage}
        <p className={styles.heading}>Attendees</p>
        <PeoplePicker
            context={this.props.context as any}
            titleText="Attendees (MOLEA)"
            placeholder='Select attendees'
            defaultSelectedUsers = {this.state.users}
            personSelectionLimit={10}
            groupName={""} // Leave this blank in case you want to filter from all users
            ensureUser={true}
            showtooltip={false}
            suggestionsLimit={5}
            required={true}
            disabled={false}
            onChange={this._getPeoplePickerItems}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
    />{attendeeFieldErrorMessage}

         {/* <p className={styles.formlabel}>Attendees (Customer)<span className={styles.required}> *</span></p> */}
         <p className={styles.formlabel}>Attendees (Customer)</p>
        <ListItemPicker listId={formconst.CONTACTS_LIST_NAME}
        context={this.props.context as any}
          columnInternalName='ContactFullName'//'EMail'
          //keyColumnInternalName='Id'
          placeholder="Select contacts"
          substringSearch={true}
          orderBy={"LastNamePhonetic"}
          itemLimit={10}
          enableDefaultSuggestions={true}
          onSelectedItem={this._onattendesSelectedItem}
          noResultsFoundText="No AttendeesattendeeDropdown Found"
          defaultSelectedItems = {(this.state.attendeeDropdown).split(',')}/>
          {/* defaultSelectedItems = {[]}/>{attcustFieldErrorMessage} */}

          <Stack horizontal verticalAlign="end" className={styles.attendeesotherstackContainer }>
          <TextField
            label="Attendees (Other)"
            value={this.state.attendeesother}
            styles={textFieldStyles as IStyleFunctionOrObject<ITextFieldStyleProps, ITextFieldStyles>}
            onChange={this.onchangeattendeesother}
          
          />
          <PrimaryButton text="+" onClick={this.handleAddattendee} />
        </Stack>
        <div>
          {interestedPartiesexternal.map((party: any, index: React.Key) => (
            <span key={index}>{party}{index !== interestedPartiesexternal.length - 1 && '; '}</span>
          ))}
        </div>    
        <br/>
        {EmailFieldErrorMessage}
        <p className={styles.heading}>Meeting Details</p>    
         <p className={styles.formlabel}>Purpose of Meeting<span className={styles.required}> *</span></p>
         <ReactQuill theme='snow'
          modules={formconst.modules}    
          formats={formconst.formats}  
          value={this.state.purposeofmeeting}  onChange={(text)=>this.onpurposeofmeetingchange(text)}  
      ></ReactQuill> {pmFieldErrorMessage}
         <div id = "purposeofmeeting" className="mt-5 text-center">
        <label htmlFor="purposeofmeetingattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="purposeofmeetingattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.purposeofmeetinghandleFileUpload}
        />

        <p id="purposeofmeetingfiles-area">
          <span id="purposeofmeetingfilesList">
            <span ref={this.filesNamesRef} id="purposeofmeetingfiles-names"></span>
          </span>
        </p>
      </div>

      <p className={styles.formlabel}>Management Summary</p>
         <ReactQuill theme='snow'
          modules={formconst.modules}    
          formats={formconst.formats}  
          value={this.state.managementsummary}  onChange={(text)=>this.onmanagementsummarychange(text)}  
      ></ReactQuill> 
       <div id = "managementsummary" className="mt-5 text-center">
        <label htmlFor="managementsummaryattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="managementsummaryattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.managementsummaryhandleFileUpload}
        />

        <p id="managementsummaryfiles-area">
          <span id="managementsummaryfilesList">
            <span ref={this.filesNamesRef} id="managementsummaryfiles-names"></span>
          </span>
        </p>
      </div>

      <p className={styles.formlabel}>Main Minutes<span className={styles.required}> *</span></p>
         <ReactQuill theme='snow'
          modules={formconst.modules}    
          formats={formconst.formats}  
          value={this.state.mainminutes}  onChange={(text)=>this.onmainminuteschange(text)}  
      ></ReactQuill> {mmFieldErrorMessage}
       <div id = "mainminutes" className="mt-5 text-center">
        <label htmlFor="mainminutesattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="mainminutesattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.mainminuteshandleFileUpload}
        />

        <p id="mainminutesfiles-area">
          <span id="mainminutesfilesList">
            <span ref={this.filesNamesRef} id="mainminutesfiles-names"></span>
          </span>
        </p>
      </div>

      <p className={styles.formlabel}>Actions</p>
         <ReactQuill theme='snow'
          modules={formconst.modules}    
          formats={formconst.formats}  
          value={this.state.actions}  onChange={(text)=>this.onactionschange(text)}  
      ></ReactQuill></div>
      <br />
      <Stack horizontal horizontalAlign='end' className={styles.stackContainer}>     
      <PrimaryButton text={buttontext} onClick={() => this._createItem(this.props)} disabled= {isbuttondisbled}/>
      <PrimaryButton text="Cancel"  onClick ={this.cancel}/>
   
      </Stack> 
      <br />
      <div>      
      {isLoading && <Spinner label="Saving, please wait..." size={SpinnerSize.large} />}
      </div>
      {imageFieldErrorMessage}
      <br />
      {FormFieldErrorMessage}
      {allRichtextsizebinaryErrorMessage}
      {successMessage}
      {failureMessage}
      </section>

    );
  }

}
