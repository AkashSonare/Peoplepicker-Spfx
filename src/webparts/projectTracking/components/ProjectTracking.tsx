import * as React from 'react';
import styles from './ProjectTracking.module.scss'
import { IProjectTrackingProps } from './IProjectTrackingProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPnPPeoplePickerState } from './IPnPPeoplePickerState';
import { sp, Web } from '@pnp/sp';
import { IButtonProps, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import * as $ from 'jquery';
import { autobind, DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react';
import { getGUID } from "@pnp/common";
import { ThemeProvider } from '@microsoft/sp-component-base';
import { mynewnumber, MYchoices} from '../../../models';
// declare var $:any;
// Import button component      
//var moment = require('moment');
import * as moment from 'moment';
const DayPickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker',

  isRequiredErrorMessage: 'Start date is required.',

  invalidInputErrorMessage: 'Invalid date format.'
};

var listItems="";
var listItems1="";
var listItems2="";
var listItems5="";

var search = window.location.search;
var params = new URLSearchParams(search);
var ProjectnameINd = params.get('ProjectName');
var managerlistid: number;
export default class ProjectTracking extends React.Component<IProjectTrackingProps, IPnPPeoplePickerState> {
  //private _opchoices: MYchoices[]=[];
  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {
      sp.setup({
        sp: {
          headers: {
            "Accept": "application/json; odata=nometadata"
          }
        }
      });
      resolve();
    });
  }
  constructor(props: IProjectTrackingProps, state: IPnPPeoplePickerState) {
    super(props);

    this.state = {
      addUsers: [],
      Projectstatus: [{ mymilestone: "", myscore: "", id: "" }],
      defaultmyusers: [],
      firstDayOfWeek: DayOfWeek.Sunday,
      value: null,
      overstatus: [{ Choices: "" }],
      DomainStatus: [{ Choices: "" }],
      ScheduleStatus: [{ Choices: "" }],
      RiskStatus: [{ Choices: "" }],
    };
    this.handleSubmit = this.handleSubmit.bind(this);
    this.fetchdatas = this.fetchdatas.bind(this);
    this.updateprojectdetails = this.updateprojectdetails.bind(this);
   // this.addSelectedUsers = this.addSelectedUsers.bind(this);
 // this.mydata();
    if (ProjectnameINd)
      this.fetchdatas();
  }
  mybutton() {
    if (!ProjectnameINd) {
      return <DefaultButton
        data-automation-id="addSelectedUsers"
        title="Add Selected Users"
        onClick={this.addSelectedUsers}> 
        Add Project
      </DefaultButton>
    } else {

      return <DefaultButton
        data-automation-id="updateSelectedUsers"
        title="Add Selected Users"
        onClick={() => this.updateSelectedUsers(ProjectnameINd)}>
        Update Project
    </DefaultButton>
    }

  }
  createUI() {


    return this.state.Projectstatus.map((el, i) => (

      <div key={i}>
        <input mycustomattribute={el.id || ''} placeholder="MileStone" name="mymilestone" value={el.mymilestone || ''} onChange={this.handleChange.bind(this, i)} />
        <input mycustomattribute={el.id || ''} placeholder="Score" name="myscore" value={el.myscore || ''} onChange={this.handleChange.bind(this, i)} />
        <input mycustomattribute={el.id || ''} type='button' value='remove' onClick={this.removeClick.bind(this, i, el.id)} />
      </div>
    ))
  }
  handleChange(i, e) {
    const { name, value } = e.target;
    let Projectstatus = [...this.state.Projectstatus];
    Projectstatus[i] = { ...Projectstatus[i], [name]: value };
    this.setState({ Projectstatus });
  }

  addClick() {
    this.setState(prevState => ({
      Projectstatus: [...prevState.Projectstatus, { mymilestone: "", myscore: "", id: "" }]
    }))
  }
  removeClick(i, delid) {
    // var delid=Number($(this).attr("ProjectStatus"));
    if (ProjectnameINd) {
      let list = sp.web.lists.getByTitle("ProjectStatus");

      list.items.getById(delid).delete().then(_ => {

        console.log("Deleted")
      });
    }
    let Projectstatus = [...this.state.Projectstatus];
    Projectstatus.splice(i, 1);
    this.setState({ Projectstatus });
  }

  async handleSubmit(event) {
    // alert('A name was submitted: ' + JSON.stringify(this.state.Projectstatus));
    //console.log(this.state.Projectstatus);
    var mystatecopy = [];
    mystatecopy.push(this.state.Projectstatus);
    //console.log(mystatecopy);
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    const batch = web.createBatch();

    const list = web.lists.getByTitle("ProjectStatus");
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    for (let k = 0; k < mystatecopy[0].length; k++) {

      list.items.inBatch(batch).add({
        ProjectName: $("#ProjectName").val(),
        MileStone: mystatecopy[0][k].mymilestone,
        Score: mystatecopy[0][k].myscore
      }, entityTypeFullName).then(b => {
        // console.log(b);
      });

    }
    batch.execute().then(() => console.log("All done!"));


    event.preventDefault();
  }
  async updateprojectdetails() {
    // alert('A name was submitted: ' + JSON.stringify(this.state.Projectstatus));
    //console.log(this.state.Projectstatus);
    var mystatecopy = [];
    mystatecopy.push(this.state.Projectstatus);
    //console.log(mystatecopy);
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    const batch = web.createBatch();

    const list = web.lists.getByTitle("ProjectStatus");
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    for (let k = 0; k < mystatecopy[0].length; k++) {
      if (!mystatecopy[0][k].id) {
        list.items.inBatch(batch).add({
          ProjectName: $("#ProjectName").val(),
          MileStone: mystatecopy[0][k].mymilestone,
          Score: mystatecopy[0][k].myscore
        }, entityTypeFullName).then(b => {
          // console.log(b);
        });
      }
      else {
        list.items.inBatch(batch).getById(mystatecopy[0][k].id).update({
          ProjectName: $("#ProjectName").val(),
          MileStone: mystatecopy[0][k].mymilestone,
          Score: mystatecopy[0][k].myscore
        }).then(b => {
          // console.log(b);
        });

      }
    }
    batch.execute().then(() => console.log("All done!"));


    // event.preventDefault();
  }


  async fetchdatas() {

    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    //const batch1 = web.createBatch();

    const list1 = web.lists.getByTitle("ProjectMaster");
    const list2 = web.lists.getByTitle("ProjectStatus");
    var myemail = [];
    var mynewid = [];
    var FetchProjectDetails = [];
    await list1.items.select('Id,ProjectManager/Id,ProjectManager/Title,ProjectManager/Name,ProjectManager/EMail,ProjectName,ProjectDetails').expand('ProjectManager').filter("ProjectName eq '" + ProjectnameINd + "'").get().then(r => {
      console.log(r)
      managerlistid = r[0].Id;
      for (let i = 0; i < r[0].ProjectManager.length; i++) {
        myemail.push(r[0].ProjectManager[i].EMail);
        mynewid.push({
          id: r[0].ProjectManager[i].Id,
          managername: r[0].ProjectManager[i].Name
        });
      }
      this.setState({ defaultmyusers: myemail });
      this.setState({ addUsers: mynewid });
      $("#ProjectName").val(r[0].ProjectName);
      $("#ProjectDetails").val(r[0].ProjectDetails);
    });

    await list2.items.select('Id,MileStone,ProjectName,Score').filter("ProjectName eq '" + ProjectnameINd + "'").top(5000).get().then(r => {
      for (let i = 0; i < r.length; i++) {
        FetchProjectDetails.push({
          mymilestone: r[i].MileStone,
          myscore: r[i].Score,
          id: r[i].Id
        })
      }
      console.log(r)
      this.setState({ Projectstatus: FetchProjectDetails });

    });
    //batch1.execute().then(() => console.log("All done!"));



  }
  public render(): React.ReactElement<IProjectTrackingProps> {
    const { firstDayOfWeek, value } = this.state;
   /* const renObjData = this.props.mychoices["Choices"].map(function(data, idx) {
      return <p key={idx}>{data}</p>;
  });*/
  var  mychoice1="";
  var  mychoice2="";
  var  mychoice3="";
  var  mychoice4="";

  if(this.props.mychoices["Choices"]){
   mychoice1= this.props.mychoices["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); }
    if(this.props.mychoices2["Choices"]){
     mychoice2= this.props.mychoices2["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); }
    if(this.props.mychoices3["Choices"]){
     mychoice3 = this.props.mychoices3["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); }
    if(this.props.mychoices4["Choices"]){
     mychoice4 = this.props.mychoices4["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); }
    return (<div>
      <div>
        <label for="ProjectName">Project Name:</label>
        <input type="text" id="ProjectName"></input>
      </div>
      <div>
        <label for="ProjectDetails">Project Scope:</label>
        <input type="text" id="ProjectDetails"></input>
      </div>
      <div>
        <label for="Highlights">Highlights:</label>
        <input type="text" id="Highlights"></input>
      </div>
      <div>
        <label for="Lowlights">Lowlights:</label>
        <input type="text" id="Lowlights"></input>
      </div>
      <div>
        <label for="Executive">Executive Status:</label>
        <input type="text" id="Executive"></input>
      </div>
      <div>
        <label for="Overall">Overall Status:</label>
        <select id="Overall">
         
        {mychoice4}
        </select>
      </div>
      <div>
        <label for="Domain">Domain Status:</label>
        <select id="Domain">
        {mychoice3}
        </select>
      </div>
      <div>
        <label for="Schedule">Schedule Status:</label>
        <select id="Schedule">
        {mychoice2}
        </select>
      </div>
      <div>
        <label for="Risk">Risk Status:</label>
        <select id="Risk">
        {mychoice1}
        </select>
      </div>
      <div>
        <label for="Charter">Charter:</label>
        <input type="text" id="Charter"></input>
      </div>
      <div>
        <label for="RoadMap">RoadMap:</label>
        <input type="text" id="RoadMap"></input>
      </div>
      <div className="docs-DatePickerExample">
        <DatePicker
          label="Target_go-Live"
          today={new Date()}
          isRequired={false}
          allowTextInput={true}
          firstDayOfWeek={firstDayOfWeek}
          strings={DayPickerStrings}
          value={value!}
          onSelectDate={this._onSelectDate}
          formatDate={this._onFormatDate}
          parseDateFromString={this._onParseDateFromString}
        />


      </div>

      <div id="ProjectManagerId">
        <PeoplePicker
          context={this.props.context}
          titleText="Project Manager"
          personSelectionLimit={3}
          groupName={""} // Leave this blank in case you want to filter from all users    
          showtooltip={true}
          isRequired={true}
          disabled={false}
          ensureUser={true}
          selectedItems={this._getPeoplePickerItems.bind(this)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          defaultSelectedUsers={this.state.defaultmyusers}
          resolveDelay={1000} />

      </div>
      <div id="OPSTeamId">
        <PeoplePicker
          context={this.props.context}
          titleText="OPS Team"
          personSelectionLimit={3}
          groupName={""} // Leave this blank in case you want to filter from all users    
          showtooltip={true}
          isRequired={true}
          disabled={false}
          ensureUser={true}
          selectedItems={this._getPeoplePickerItems.bind(this)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          defaultSelectedUsers={this.state.defaultmyusers}
          resolveDelay={1000} />

      </div>
      <div id="OCMTeamId">
        <PeoplePicker
          context={this.props.context}
          titleText="OCM Team"
          personSelectionLimit={3}
          groupName={""} // Leave this blank in case you want to filter from all users    
          showtooltip={true}
          isRequired={true}
          disabled={false}
          ensureUser={true}
          selectedItems={this._getPeoplePickerItems.bind(this)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          defaultSelectedUsers={this.state.defaultmyusers}
          resolveDelay={1000} />

      </div>
      <div id="Others">
        <PeoplePicker
          context={this.props.context}
          titleText="Others"
          personSelectionLimit={3}
          groupName={""} // Leave this blank in case you want to filter from all users    
          showtooltip={true}
          isRequired={true}
          disabled={false}
          ensureUser={true}
          selectedItems={this._getPeoplePickerItems.bind(this)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          defaultSelectedUsers={this.state.defaultmyusers}
          resolveDelay={1000} />

      </div>
      <div id="Sponsor">
        <PeoplePicker
          context={this.props.context}
          titleText="Sponsor"
          personSelectionLimit={3}
          groupName={""} // Leave this blank in case you want to filter from all users    
          showtooltip={true}
          isRequired={true}
          disabled={false}
          ensureUser={true}
          selectedItems={this._getPeoplePickerItems.bind(this)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          defaultSelectedUsers={this.state.defaultmyusers}
          resolveDelay={1000} />

      </div>
{this.mybutton()}
      
      {this.createUI()}
      <a href="#" id="Addbutton" onClick={this.addClick.bind(this)}>Addbtn</a><br></br>


    </div>

    );

  }


  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
    this.setState({ addUsers: items });
  }
  @autobind 
  private addSelectedUsers(): void{
    this.handleSubmit(event);

    var peoplepicarray = [];
    for (let i = 0; i < this.state.addUsers.length; i++) {

      peoplepicarray.push(this.state.addUsers[i]["id"]);

    }

    let projectname = $("#ProjectName").val();
    let projectdetail = $("#ProjectDetails").val();
    let TargetGolive = this.state.value;
    var TargetGolive1 = moment(TargetGolive, 'DD.MM.YYYY');
    TargetGolive1.toISOString();
    var AlbumData ={
      "ProgramName": projectname,
      "ProgramManagerId": { "results": peoplepicarray },
      //"ProjectDetails":projectdetail,
      "ProjectScope": projectdetail,
      "Highlights": $("#Highlights").val(),
      "Lowlights": $("#Lowlights").val(),
      "ExecutiveStatus": $("#Executive").val(),
      "TargetGoLive": TargetGolive1,
      Charter: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "Description",
        Url: "https://5.imimg.com/data5/AA/KK/MY-6677193/red-rose-500x500.jpg"
      },
      RoadMap: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "Description",
        Url: "https://5.imimg.com/data5/AA/KK/MY-6677193/red-rose-500x500.jpg"
      },
     // "ProgramManagerId": { "results": peoplepicarray },
      "OPSTeamId": { "results": peoplepicarray },
      "OCSTeamId": { "results": peoplepicarray },
      "OthersId": { "results": peoplepicarray },
      "SponsorId": { "results": peoplepicarray },
      "OverAllStatus": $("#Overall option:selected").val(),
      "DomainStatus": $("#Domain option:selected").val(),
      "ScheduleStatus": $("#Schedule option:selected").val(),
      "RiskStatus": $("#Risk option:selected").val()
     
      /*  ProjectManager: {   
            results: this.state.addUsers  
        }  */
    };
    sp.web.lists.getByTitle("Program").items.add(AlbumData).then(i => {
      console.log(i);
    });

  }
  @autobind
  private updateSelectedUsers(ProjectnameINd): void {
    this.updateprojectdetails();

    var peoplepicarray = [];
    for (let i = 0; i < this.state.addUsers.length; i++) {

      peoplepicarray.push(this.state.addUsers[i]["id"]);

    }

    let projectname = $("#ProjectName").val();
    let projectdetail = $("#ProjectDetails").val();
    sp.web.lists.getByTitle("ProjectMaster").items.getById(managerlistid).update({
      "ProjectName": projectname,
      "ProjectManagerId": { "results": peoplepicarray },
      "ProjectDetails": projectdetail
      /*  ProjectManager: {   
            results: this.state.addUsers  
        }  */
    }).then(i => {
      console.log(i);
    });

  }
  private _onSelectDate = (date: Date | null | undefined): void => {
    this.setState({ value: date });
  };


  private _onFormatDate = (date: Date): string => {
    let mm: number | string = date.getMonth() + 1;
    mm = (mm < 10) ? '0' + mm : mm;
    return date.getDate() + '/' + mm + '/' + (date.getFullYear());
  };

  private _onParseDateFromString = (value: string): Date => {
    const date = this.state.value || new Date();
    const values = (value || '').trim().split('/');
    const day = values.length > 0 ? Math.max(1, Math.min(31, parseInt(values[0], 10))) : date.getDate();
    const month = values.length > 1 ? Math.max(1, Math.min(12, parseInt(values[1], 10))) - 1 : date.getMonth();
    let year = values.length > 2 ? parseInt(values[2], 10) : date.getFullYear();
    if (year < 100) {
      year += date.getFullYear() - (date.getFullYear() % 100);
    }
    return new Date(year, month, day);
  };
  /*private onaddbtns = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  
    this.props.onAddButton();
  }  
  private onUpdatebtns = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  
    this.props.onDeleteBtn();
  }*/

}
