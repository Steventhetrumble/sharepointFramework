import { ListResult, ListRow } from "../../common/IObjects";
import { SharepointRestProvider } from "../../dataproviders/SharepointRestProvider";

export interface ITargetedSitesState {
  showPanel: boolean;
  siteIsSelected: boolean;
  siteSelected: ListRow;
  showDeletionDialog: boolean;
}

export interface ITargetedSitesProps {
  totalRows: number;
  targets: ListResult;
  provider: SharepointRestProvider;
  onSiteDeletion(): void;
}
