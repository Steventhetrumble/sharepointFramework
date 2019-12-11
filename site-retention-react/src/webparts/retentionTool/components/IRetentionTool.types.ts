import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISPSite } from "./common/IObjects";
import { SharepointRestProvider } from "./dataproviders/SharepointRestProvider";

export interface IRetentionToolProps {
  description: string;
  context: WebPartContext;
}

export interface IRetentionToolState {
  showPivot: boolean;
  sitesPromise: Array<Promise<ISPSite>>;
  provider: SharepointRestProvider;
}
