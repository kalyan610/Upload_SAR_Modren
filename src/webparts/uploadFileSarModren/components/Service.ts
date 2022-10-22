import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";


export default class Service {

    public mysitecontext: any;

    public constructor(siteUrl: string, Sitecontext: any) {
        this.mysitecontext = Sitecontext;


        sp.setup({
            sp: {
                baseUrl: siteUrl

            },
        });

    }



    public async GetAllQuarters():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("SAR_Quarter").items.select('Title','ID').expand().get().then(function (data) {
 
     return data;
 
 
     });
 
 
    }

    public async GetAllApplications():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("Applications").items.select('Title','ID').expand().get().then(function (data) {
 
     return data;
 
 
     });
 
 
    }

    public async uploadFile(fileDetails:any,data:any){

        try
        {
   
console.log(this.mysitecontext);
     
   
   const file = await sp.web.getFolderByServerRelativeUrl(this.mysitecontext.pageContext.web.serverRelativeUrl + "/SAR_Docs").files.add(fileDetails.name, fileDetails, true);
   const item = await file.file.getItem();
   
   await item.update(data);
   
   alert("File uploaded sucessfully");
   window.location.replace("https://capcoinc.sharepoint.com/sites/SARModren_Dev/SitePages/Upload-Form.aspx");
         
   
   
       }
   
       catch(error){
           console.log(error);
       }
   
   }

}