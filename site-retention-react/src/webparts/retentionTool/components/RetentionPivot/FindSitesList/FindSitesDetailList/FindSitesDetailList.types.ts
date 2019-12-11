import { IColumn } from "office-ui-fabric-react";
import { ISPSite, IList, ListResult } from "../../../common/IObjects";

export interface ListRows {
  value: ListRow[];
}

export interface ListRow {
  Title: string;
}

export interface IFindSitesDetailListState {
  columns: IColumn[];
  items: ISPSite[];
  selectionDetails: string;
  announcedMessage?: string;
  lists: IList[];
  targetedSite: ISPSite;
  targetSiteSelected: boolean;
  sitesAlreadyTargeted: TargetedSites;
}

export interface IFindSitesDetailListProps {
  items: ISPSite[];
  siteSelectedCallback(site: ISPSite): void;
  siteUnselectedCallback(): void;
  alreadyTargetd: ListResult;
}

export interface TargetedSites {
  [key: string]: boolean;
}
