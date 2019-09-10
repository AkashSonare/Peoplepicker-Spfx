import { autobind,DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react'; 
export interface IPnPPeoplePickerState {  
    addUsers: string[];  
    Projectstatus?:any[];
    defaultmyusers?:any[];
    firstDayOfWeek?: DayOfWeek;
    value?: Date | null;
    overstatus: any[];
    DomainStatus: any[];
    ScheduleStatus:any[];
    RiskStatus: any[];
} 