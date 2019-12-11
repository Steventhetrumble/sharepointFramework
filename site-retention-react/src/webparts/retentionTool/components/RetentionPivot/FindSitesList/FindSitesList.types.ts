import {
  ISPSite,
  ListResult,
  SPGroupList,
  ISPUsers,
  ISPUser
} from "../../common/IObjects";
import { SharepointRestProvider } from "../../dataproviders/SharepointRestProvider";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISharepointRestProvider } from "../../dataproviders/ISharepointRestProvider.types";
import { IDropdownOption } from "office-ui-fabric-react";

export interface IFindSitesListState {
  showPanel: boolean;
  showUserModal: boolean;
  siteIsSelected: boolean;
  siteSelected: ISPSite;
  subWebs: ISPSite[];
  groups: SPGroupList;
  displayOptions: IDropdownOption[];
  allGroups: { [id: number]: ISPUsers };
  idsSelected: number[];
  users: ISPUser[];
}

export interface IFindSitesListProps {
  totalRows: number;
  items: ISPSite[];
  targets: ListResult;
  provider: ISharepointRestProvider;
  onSiteAddition(): void;
}
