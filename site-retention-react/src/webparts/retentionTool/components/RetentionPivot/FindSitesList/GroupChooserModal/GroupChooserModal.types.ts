import {
  ISPSite,
  SPGroupList,
  ISPUsers,
  ISPUser
} from "../../../common/IObjects";
import { ISharepointRestProvider } from "../../../dataproviders/ISharepointRestProvider.types";
import { IDropdownOption } from "office-ui-fabric-react";

export interface IGroupChooserModalState {}

export interface IGroupChooserModalProps {
  groups: SPGroupList;
  siteSelected: ISPSite;
  showModal: boolean;
  onClickCancel(): void;
  onClickAddTarget(): void;
  onChangeSelectedIds(ids: number[]): void;
  provider: ISharepointRestProvider;
  displayOptions: IDropdownOption[];
  allGroups: { [id: number]: ISPUsers };
  selectedIds: number[];
  users: ISPUser[];
}
