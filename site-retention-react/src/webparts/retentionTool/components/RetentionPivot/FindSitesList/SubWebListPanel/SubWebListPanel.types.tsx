import { ISPSite } from "../../../common/IObjects";

export interface ISubWebListPanelState {
    
  }
  
export interface ISubWebListPanelProps {
    subSites: ISPSite[];
    showPanel: boolean;
    onClickClose(): void;
    onClickAddSite(): void;
}