import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { ISPSite, ISPDate } from '../../../common/IObjects';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import {
    Result,
    RelevantResults,
    Row
} from '../../IManageLists';
import { SPSearchPanelSubSiteList } from './SPSearchPanelSubSiteList/SPSearchPanelSubSiteList';
import {  SPModal } from './SPPeopleDropdown/SPPeopleDropDown';
import { SPGroupList, SPGroup } from '../../../common/IObjects';
import {TargetedSites} from '../SPDetailList';
import { Label } from 'office-ui-fabric-react';

export interface ISPSearchPanelState {
    showPanel: boolean;
    targetedSite: ISPSite;
    subWebs: ISPSite[];
    updated: boolean;
    showModal: boolean;
    userGroups: SPGroupList;
    
}
export interface ISPSearchPanelProps {
    showPanel: boolean;
    targetedSite: ISPSite;
    spHttpClient: SPHttpClient;
    sitesAlreadyTargeted: TargetedSites;
    siteAddedCallback(s: string): void;
}

const warningStyle = {
    root: {
      color: 'red'
    }
  };

export class SPSearchPanel extends React.Component<ISPSearchPanelProps, ISPSearchPanelState> {
    private _allItems: ISPSite[];
    

    constructor(props: ISPSearchPanelProps) {
        super(props);

        this.state = {
            showPanel: false,
            targetedSite: null,
            subWebs: [],
            updated: false,
            showModal: false,
            userGroups: null
        };
    }

    public componentDidUpdate(previousProps: any, previousState: ISPSearchPanelState) {
        if(previousState.showModal == false && this.state.showModal == true){
            this.setState({
                showModal: false
            });
        }
        if (this.props.targetedSite != null && this.state.targetedSite != this.props.targetedSite && this.state.showPanel == true) {

            this.setState({
                targetedSite: this.props.targetedSite
            });
            this._getSubWebs().then((res: RelevantResults) => {
                const randomDate = _randomDate(new Date(2012, 0, 1), new Date(2012, 0, 1));
                let _sites: ISPSite[] = [];
                res.Table.Rows.forEach((row: Row) => {
                    this._getSPDate(row).then((responseDate: ISPDate) => {
                        _sites.push({
                            DocID: row.Cells[1].Value,
                            Title: row.Cells[2].Value,
                            Url: row.Cells[3].Value,
                            ViewsLifeTime: Number(row.Cells[6].Value),
                            ViewsRecent: Number(row.Cells[7].Value),
                            Size: Number(row.Cells[8].Value),
                            SiteDescription: row.Cells[9].Value,
                            LastItemUserModifiedDateSharepoint: (responseDate.error == null) ? responseDate.value : "randomDate.date",
                            LastItemUserModifiedDate: (responseDate.error == null) ? responseDate.date : randomDate.date,
                            LastItemUserModifiedDatevalue: (responseDate.error == null) ? responseDate.datevalue : randomDate.value,
                            LastItemUserModifiedDateFomatted: (responseDate.error == null) ? responseDate.dateFormatted : randomDate.dateFormatted,
                            renderTemplateId: row.Cells[16].Value
                        });
                    });
                });
                this._allItems = _sites;
                this.setState({
                    subWebs: this._allItems,
                    updated: true
                });
            });
        }
    }

    private _getSubWebs(): Promise<RelevantResults> {
        return this.props.spHttpClient.get(this.props.targetedSite.Url + `/_api/search/query?querytext=%27(contentclass:STS_Web) Path:${this.props.targetedSite.Url}/* NOT (WebTemplate:GROUP)%27&trimduplicates=false&selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate,WebTemplate%27`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json().then((responseJSON: Result) => {
                    return responseJSON.PrimaryQueryResult.RelevantResults;
                });
            });
    }
    private _getUserGroups(): Promise<SPGroupList> {
        return this.props.spHttpClient.get(this.props.targetedSite.Url + `/_api/Web/SiteGroups`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json().then((responseJSON: SPGroupList) => {
                    console.log(responseJSON);
                    return responseJSON;
                });
            });
    }

    private _convertSPDate(date: string): Date {

        let xDate: string = date.split("T")[0];
        let xTime: string = date.split("T")[1];
        xTime = xTime.split("Z")[0];

        // split apart the hour, minute, & second
        let xTimeParts: string[] = xTime.split(":");
        let xHour: string = xTimeParts[0];
        let xMin: string = xTimeParts[1];
        let xSec: string = xTimeParts[2];

        // split apart the year, month, & day
        let xDateParts: string[] = xDate.split("-");
        let xYear: string = xDateParts[0];
        let xMonth: string = xDateParts[1];
        let xDay: string = xDateParts[2];

        // REALLY STRANGE ----- subtract 1 from month because it starts at zero ie 0 == january
        let dDate: Date = new Date(Number(xYear), Number(xMonth) - 1, Number(xDay), Number(xHour), Number(xMin), Number(xSec));
        return dDate;
    }

    private _getSPDate(row: Row): Promise<ISPDate> {
        return this.props.spHttpClient.get(row.Cells[3].Value + `/_api/Web/LastItemUserModifiedDate`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                if (response.status != 200) {
                    console.log(response.status);
                    console.log(response.statusText);
                }
                return response.json().then((responseJSON: ISPDate) => {

                    if (responseJSON.error == null) {
                        let _date: Date = this._convertSPDate(responseJSON.value);
                        return {
                            value: responseJSON.value,
                            dateFormatted: _date.toLocaleDateString(),
                            datevalue: _date.valueOf(),
                            error: null
                        };
                    }
                    else {
                        return {
                            value: null,
                            DateFormatted: null,
                            datevalue: null,
                            error: responseJSON.error
                        };
                    }
                }).catch(err => {
                    console.log(`convert to json error in SPDATE: ${err}`);
                    return err;
                });
            }).catch(err => {
                console.log(`get request error at SPDATE: ${err}`);
                return err;
            });
    }

    public render() {
        return (
            <div>
                <DefaultButton secondaryText="Opens the Sample Panel" onClick={this._showPanel} text="Open Panel" />
                <Label styles={warningStyle}>{(this.props.showPanel && this.props.sitesAlreadyTargeted[this.props.targetedSite.Title])? `${this.props.targetedSite.Title} is already targeted for deletion`: ''}</Label>
                <Panel
                    isOpen={this.state.showPanel && this.props.showPanel}
                    type={PanelType.medium}
                    onDismiss={this._hidePanel}
                    headerText={`Target Site for Deletion: `}
                    closeButtonAriaLabel="Close"
                    onRenderFooterContent={this._onRenderFooterContent}
                >
                <Label styles={warningStyle}>{(this.props.showPanel && this.props.sitesAlreadyTargeted[this.props.targetedSite.Title])? `${this.props.targetedSite.Title} is already targeted for deletion`: ''}</Label>
                <SPSearchPanelSubSiteList sites={this.state.subWebs} updated={this.state.updated} ></SPSearchPanelSubSiteList>
                </Panel>
                <SPModal 
                showModal={this.state.showModal} 
                site={this.props.targetedSite} 
                spHttpClient={this.props.spHttpClient} 
                groups={this.state.userGroups}
                sitesAlreadyTargeted={this.props.sitesAlreadyTargeted}
                siteAddedCallback={this.props.siteAddedCallback} 
                />
            </div>
        );
    }

    private _onRenderFooterContent = () => {
        return (
            <div>
                <PrimaryButton onClick={
                    () => {
                        this._addSite();
                    }
                } style={{ marginRight: '8px' }}>
                    Add Site
                </PrimaryButton>
                <DefaultButton onClick={this._hidePanel}>Cancel</DefaultButton>
            </div>
        );
    }

    

    private _showPanel = () => {
        this.setState({
            showPanel: true

        });
    }

    private _hidePanel = () => {
        this.setState({
            showPanel: false,
            targetedSite: null,
            subWebs: [],
            updated: false,
            showModal: false,
            userGroups: null

        });
    }
    private _addSite = () => {
        this._getUserGroups().then((groups: SPGroupList) => {
            this.setState({
                userGroups: groups,
                showModal: true
            });
            this._hidePanel();
        });
    }
}

function _randomDate(start: Date, end: Date): { value: number; dateFormatted: string, date: Date } {
    const _date: Date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
    return {
        date: _date,
        value: _date.valueOf(),
        dateFormatted: _date.toLocaleDateString()
    };
}