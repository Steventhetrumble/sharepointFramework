import {
  IList,
  ISPSite,
  Row,
  ListResult,
  ISPDate,
  ListRow,
  RelevantResults,
  SPGroupList,
  ISPUsers,
  ISPUser
} from "./../common/IObjects";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";

export interface ISharepointRestProvider {
  deleteSiteFromList(siteSelected: ListRow): void;
  getAllLists(): Promise<IList[]>;
  getRootSiteData(): Promise<ISPSite>;
  getSitesListing(rowStart: number): Promise<Row[]>;
  getTargetedSitesListing(rowStart: number): Promise<ListResult>;
  getSPDate(row: Row): Promise<ISPDate>;
  getSubWebs(target: ISPSite): Promise<RelevantResults>;
  getUserGroups(target: ISPSite): Promise<SPGroupList>;
  lookupDates(res: RelevantResults): Promise<ISPSite>[];
  makeOptions(
    groups: SPGroupList,
    url: string
  ): Promise<[IDropdownOption[], { [id: number]: ISPUsers }]>;
  getByEmail(email: string): void;
  postListData(site: ISPSite, users: ISPUser[]): void;
  saveTarget(targetedSite: ListRow): void;
  saveEmail(contents: string): void;
  getEmails(): Promise<string[]>;
  // getAllSites(): Promise<Array<Promise<ISPSite>>>;
}
