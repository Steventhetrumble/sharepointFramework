import { ISPSite, ListResult } from "../common/IObjects";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SharepointRestProvider } from "../dataproviders/SharepointRestProvider";

export interface IRetentionPivotProps {
  context: WebPartContext;
  provider: SharepointRestProvider;
}

export interface IRetentionPivotState {
  items: ISPSite[];
  totalRows: number;
  targets: ListResult;
  targetRows: number;
  emails: string[];
}
