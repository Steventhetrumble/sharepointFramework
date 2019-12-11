import { IColumn } from "office-ui-fabric-react";
import { ISPSite, IList, ListResult, ListRow } from "../../../common/IObjects";

export interface ListRows {
  value: ListRow[];
}



export interface ITargetedSitesDetailListState {
  columns: IColumn[];
  items: ListResult;
  selectionDetails: string;
  announcedMessage?: string;
  lists: IList[];
  targetedSite: ListRow;
  targetSiteSelected: boolean;
  sitesAlreadyTargeted: TargetedSites;
}

export interface ITargetedSitesDetailListProps {
  items: ListResult;
  siteSelectedCallback(site: ListRow): void;
  siteUnselectedCallback(): void;
}

export interface TargetedSites {
  [key: string]: boolean;
}
