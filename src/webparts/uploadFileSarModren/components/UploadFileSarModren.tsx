import * as React from 'react';
import styles from './UploadFileSarModren.module.scss';
import { IUploadFileSarModrenProps } from './IUploadFileSarModrenProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ChoiceGroup,IChoiceGroupOption, textAreaProperties,Stack, IStackTokens, StackItem,IStackStyles } from 'office-ui-fabric-react'; 
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption} from 'office-ui-fabric-react/lib/Dropdown';
//#region GlobalVaraibles
import Service from './Service';

import { Button,PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';  


const sectionStackTokens: IStackTokens = { childrenGap: 10 };
const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { padding: 10} };
const stackButtonStyles: Partial<IStackStyles> = { root: { width: 20 } };
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

let RootUrl = '';

export interface SARUploadControlFieldsState{
  operation:any;
  QuarterListItems: any;
  MyQuarterValue:any;
  ApplicationListItems:any;
  MyApplicationValue:any;
  file:any;
  dtreqdate:Date;

}

export default class UploadFileSarModren extends React.Component<IUploadFileSarModrenProps, SARUploadControlFieldsState> {
  public _service: any;
  public GlobalService: any;
  protected ppl;

  public constructor(props:IUploadFileSarModrenProps) {
    super(props);
    this.state={
      
      operation:null,
      QuarterListItems: [],
      MyQuarterValue:null,
      ApplicationListItems:[],
      MyApplicationValue:null,
      file:null,
      dtreqdate:null
    };

    RootUrl = this.props.url;

this._service = new Service(this.props.url, this.props.context);

this.GlobalService = new Service(this.props.url, this.props.context);

this.GetAllQuarters();

this.GetAllApplications();


  }

  public async GetAllQuarters() {

    var myQuartersLocal: any = [];

    var data = await this._service.GetAllQuarters();

    console.log(data);

    var AllQuarters: any = [];

    for (var k in data) {

      AllQuarters.push({ key: data[k].Title, text: data[k].Title});
    }

    console.log(AllQuarters);

    
   this.setState({ QuarterListItems: AllQuarters });
  

  }

  public async GetAllApplications() {

    var myApplicationLocal: any = [];

    var data = await this._service.GetAllApplications();

    console.log(data);

    var AllApplication: any = [];

    for (var k in data) {

      AllApplication.push({ key: data[k].Title, text: data[k].Title});
    }

    console.log(AllApplication);

    
   this.setState({ ApplicationListItems: AllApplication });
  

  }



  private handleQuarter(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {

    
    this.setState({ MyQuarterValue:item.key });

    
  }

  private handleApplication(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {
    
    this.setState({ MyApplicationValue:item.key });
    
  }

  private fileChangeHandler(event:any):void{
    //this.setState({date:data});
    this.setState({file:event.target.files[0]});
    

  }

  public handleRequestDateChange = (date: any) => {

    this.setState({ dtreqdate: date });

    }
    private OnBtnClick():void{

      const mycurrentdate = new Date();

    
    if (this.state.MyQuarterValue == null || this.state.MyQuarterValue == 'Select Quarter') {

      alert('Please select Quarter');
      //this.setState({ flag: false });
     
    }

    else if(this.state.MyApplicationValue == null || this.state.MyApplicationValue == 'Select Application')
    {
     
      alert('please select Application');
    }

    else if(this.state.file==null)
    {
     
      alert('please select any file');
    }

    else if(this.state.dtreqdate==null)
    {

      alert('please select Deadline for submission');
    }

    else if(this.state.dtreqdate>mycurrentdate)
    {

      alert('Submission date should be less than current date');

    }

    

    else
    {

let date1=(this.state.dtreqdate.getDate()+1);

let month1= (this.state.dtreqdate.getMonth()+1);

let year1 =(this.state.dtreqdate.getFullYear());

let FinalRequestDelDate=month1+'/'+this.state.dtreqdate.getDate() +'/' +year1;



    
    let inputData:any=
    {
      Quarter: this.state.MyQuarterValue,
      Application:this.state.MyApplicationValue,
      Status:"Not Copied",
      Tentative_x0020_Date:FinalRequestDelDate
      
      
    };

    this._service.uploadFile(this.state.file,inputData);


  }

  

    }

 
  
  
  public render(): React.ReactElement<IUploadFileSarModrenProps> {
    return (
      <Stack tokens={stackTokens} styles={stackStyles} >
        <Stack>
        <b><label className={styles.labelsFonts}>Quarter <label className={styles.recolorss}>*</label></label></b><br/>  
        
            <Dropdown className={styles.onlyFont}
                placeholder="Select  Quarter"
                options={this.state.QuarterListItems}
                styles={dropdownStyles}
                selectedKey={this.state.MyQuarterValue ? this.state.MyQuarterValue : undefined} onChange={this.handleQuarter.bind(this)}/>
                <br></br>
              
                <b><label className={styles.labelsFonts}>Application <label className={styles.recolorss}>*</label></label></b><br/>
                
                <Dropdown className={styles.onlyFont}
                placeholder="Select  Application"
                options={this.state.ApplicationListItems}
                styles={dropdownStyles}
                selectedKey={this.state.MyApplicationValue ? this.state.MyApplicationValue : undefined} onChange={this.handleApplication.bind(this)}/>
              <br/>
              <b><label className={styles.labelsFonts}>File Name <label className={styles.recolorss}>*</label></label></b><br/>
              <input type="file" name="file" onChange={this.fileChangeHandler.bind(this)} accept=".xlsx"/>
               <br></br>
               <b><label className={styles.labelsFonts}>Submission DeadLine <label className={styles.recolorss}>*</label></label></b><br/>
             <div className={styles.boxsize}>
           <DateTimePicker  
          dateConvention={DateConvention.Date}  
          showLabels={false}
          value={this.state.dtreqdate}  
          onChange={this.handleRequestDateChange}
           />  
        </div><br>
        </br>
        <PrimaryButton text="Submit" onClick={this.OnBtnClick.bind(this)} styles={stackButtonStyles} className={styles.Mybutton}/>
        </Stack>
        </Stack>
               
    );
  }
}
