import { SPFI } from "@pnp/sp";
import { getSP } from "../pnpjsconfig";
import { IListMM, IMeetingMinutesFormProps } from "../webparts/meetingMinuteEditWebpart/components/IMeetingMinutesFormProps";
import * as formconst from "../webparts/constant";


export const getCustomerItem= async (props:IMeetingMinutesFormProps) => {
    const _sp :SPFI = getSP(props.context) ;
    return new Promise((resolve,reject) =>{
      _sp.web.lists.getByTitle(formconst.METADATA_LISTNAME).items.select("Title").orderBy("ID", false).top(1)()
      .then((items) => {
        if (items.length > 0) {
        
        console.log(items) 
          resolve(items);
        } else {
          reject(new Error("Customer not found"));
        }

    }).catch((error) => {
      reject(error);
    });
});
}
    
export const submitDataAndGetId = async (props:IMeetingMinutesFormProps,data:any): Promise<any> => {
  
  const _sp :SPFI = getSP(props.context) ;
 
  return _sp.web.lists.getByTitle(formconst.LISTNAME).items.add(data)
    .then((response) => {
      console.log(response)
     
        const itemId = response.data.Id;

        console.log("ID",itemId)
        // Resolve the promise with the item ID
        return Promise.resolve(itemId);
    })
    .catch((error) => {
        // Handle any errors that occurred during the request
        return Promise.reject(error);
    });

  
}


export const updateData=(props:IMeetingMinutesFormProps ,itemId: number, data: any): Promise<void>=> {
  const _sp :SPFI = getSP(props.context) ;
  return new Promise<void>((resolve, reject) => {
    _sp.web.lists.getByTitle(formconst.LISTNAME).items.getById(itemId).update(data)
      .then(() => {
        
        //console.log(e.response.headers.get("content-length"))
        resolve();
      })
      .catch((error) => {
 
        reject(error);
      });
  });
}

/* export const writeFile = (props:IMeetingMinutesFormProps,folderUrl:any,filename:any, file:any,meetingid:string,weburl?:any)=>{
  const _sp :SPFI = getSP(props.context) ;
  const list = _sp.web.lists.getByTitle(formconst.LIBRARYNAME);
  
  const fileContent = "Hello, world!";
  const fileName = "example.txt";

  try {
    const file = await list.rootFolder.files.add(fileName, fileContent);
    console.log(`File '${fileName}' has been written to SharePoint successfully!`);
  } catch (error) {
    console.error("Error writing to file:", error);
  }
}
 */
export const uploadAttachment = (props:IMeetingMinutesFormProps,folderUrl:any,filename:any, file:any,meetingid:string,weburl?:any)=>{
const _sp :SPFI = getSP(props.context) ;
  
_sp.web.folders.addUsingPath(folderUrl);
  return new Promise((resolve,reject) =>{
    _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(filename, file)
    .then((items) => {
      
      items.file.getItem().then((item)=>{

        item.update({MeetingId:meetingid})
      })
   
      resolve(items);
    })
  }).catch((error) => {
    reject(error);
  });



}

export const getItem= async (props:IMeetingMinutesFormProps,key:string):Promise<IListMM[]>=> {

  console.log(key)
  const _sp :SPFI = getSP(props.context) ;
  const items  =_sp.web.lists.getByTitle(formconst.LISTNAME).items.filter(`Title eq '${key}'`)()
    
  console.log('items',items);
  return items;
    
    
}

export const readFile = (props:IMeetingMinutesFormProps,filepath:string):Promise<string> =>{
 console.log(filepath)
  const _sp :SPFI = getSP(props.context) ;
  try {
      const fileContents = _sp.web.getFileByServerRelativePath(filepath).getText();
      console.log('File content:');
      console.log(fileContents);
      return fileContents
  } catch (error) {
      console.error(`Error reading the SharePoint file: ${error}`);
  }
}


function reject(error: any) {
  throw new Error("Function not implemented.");
}




  

  

   
  

 
  
  
  
  
  
  
  
  
  