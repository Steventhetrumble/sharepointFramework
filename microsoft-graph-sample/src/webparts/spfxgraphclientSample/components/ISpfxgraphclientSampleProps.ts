import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ClientMode } from "./ClientMode";

export interface ISpfxgraphclientSampleProps {
  description: string;
  context: WebPartContext;
  clientMode: ClientMode;
}
