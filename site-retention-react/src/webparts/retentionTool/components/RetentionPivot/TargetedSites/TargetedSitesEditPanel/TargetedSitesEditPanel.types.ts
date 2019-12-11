import { DayOfWeek, IDatePickerStrings } from "office-ui-fabric-react";
import { ListRow } from "../../../common/IObjects";

export interface ITargetedSitesEditPanelState {
  firstDayOfWeek?: DayOfWeek;
  dateSelected: boolean;
  showPanel: boolean;
  Title: string;
  Deleted: boolean;
  FirstEmailReply: boolean;
  FirstEmailSent: boolean;
  SecondEmailReply: boolean;
  SecondEmailSent: boolean;
  URL: string;
  UserLastModified: string;
  DeletionDate: string;
  RecentViews: boolean;
}

export interface ITargetedSitesEditPanelProps {
  siteSelected: ListRow;
  showPanel: boolean;
  onClickClose(): void;
  onClickSave( modifiedListRow: ListRow): void;
}


export const DayPickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  
    goToToday: 'Go to today',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker'
  };