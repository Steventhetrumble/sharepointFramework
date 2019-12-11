export interface IList {
  Title?: string;
  Id?: string;
}

export interface ISPUsers {
  value: ISPUser[];
  error: Error;
  message: string;
  OwnerTitle: string;
}
export interface ISPUser {
  Title: string;
  Email: string;
}
export interface ISPDate {
  date: Date;
  value: string;
  dateFormatted: string;
  datevalue: number;
  error: Error;
}

export interface ISPGroupEmails {
  groups: ISPGroupEmail[];
}

export interface ISPGroupEmail {
  GroupName: string;
  userEmails: string[];
}

export interface ISPSite {
  Title: string;
  Url: string;
  ViewsLifeTime: number;
  ViewsRecent: number;
  Size: number;
  SiteDescription: string;
  LastItemUserModifiedDateSharepoint: string;
  LastItemUserModifiedDateFomatted: string;
  LastItemUserModifiedDatevalue: number;
  LastItemUserModifiedDate: Date;
  renderTemplateId: string;
}

export interface SPGroupList {
  value: SPGroup[];
  error: Error;
}

export interface SPGroup {
  Id: number;
  OwnerTitle: string;
  users: ISPUsers;
}

export interface Usage {
  Storage: number;
  ViewsLifeTime: number;
  ViewsRecent: number;
  Size: number;
}

export interface IRequiredSymbols {
  [requiredSymbols: string]: boolean;
}

export interface IListResult {
  value: IListRow[];
}

export interface IListRow {
  Title: string;
  bodyTemplate: string;
}

export interface Result {
  PrimaryQueryResult: PrimaryQueryResult;
}

export interface PrimaryQueryResult {
  RelevantResults: RelevantResults;
}

export interface RelevantResults {
  Table: Table;
  TotalRows: number;
}

export interface Table {
  Rows: Row[];
}

export interface Row {
  Cells: Cell[];
  length: number;
}

export interface Cell {
  Key: string;
  Value: string;
}

export interface ListResult {
  value: ListRow[];
}

export interface ListRow {
  "@odata.editLink": string;
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
  emails: string;
}

export interface userID {
  value: number;
}
