import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { ListRow } from '../../SPTargetedSitesDetailList';
import { PrimaryButton, DefaultButton, Stack, IStackTokens, Label, Toggle } from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react';

export interface ISiteEditPanelProps {
    refreshCallback(): void;
    siteSelected: ListRow;
    showPanel: boolean;
    spHttpClient: SPHttpClient;
}

export interface ISiteEditPanelState {
    firstDayOfWeek?: DayOfWeek;
    dateSelected: boolean;
    showPanel: boolean;
    Title: string;
    Deleted: boolean;
    FirstEmailReply: boolean;
    FirstEmailSent: boolean;
    SecondEmailReply: boolean;
    SecondEmailSent: boolean;
    URL: string;
    UserLastModified: string;
    DeletionDate: string;
    RecentViews: boolean;
}

const stackTokens: IStackTokens = { childrenGap: 10 };

const DayPickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  
    goToToday: 'Go to today',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker'
  };

  export interface IDatePickerBasicExampleState {
    
  }
  


export class SiteEditPanel extends React.Component<ISiteEditPanelProps, ISiteEditPanelState> {

    constructor(props: ISiteEditPanelProps) {
        super(props);
        this.state = {
            firstDayOfWeek: DayOfWeek.Sunday,
            showPanel: false,
            Title: "",
            Deleted: null,
            FirstEmailReply: null,
            FirstEmailSent: null,
            SecondEmailReply: null,
            SecondEmailSent: null,
            URL: "",
            UserLastModified: null,
            DeletionDate: "",
            RecentViews: null,
            dateSelected: false
        };
    }

    public componentDidUpdate(previousProps: ISiteEditPanelProps, previousState: ISiteEditPanelState) {
        if (!previousProps.showPanel && this.props.showPanel && this.props.siteSelected != null) {
            this._showPanel();
            this.setState({
                Title: this.props.siteSelected.Title,
                Deleted: this.props.siteSelected.Deleted,
                FirstEmailReply: this.props.siteSelected.FirstEmailReply,
                FirstEmailSent: this.props.siteSelected.FirstEmailSent,
                SecondEmailReply: this.props.siteSelected.SecondEmailSent,
                SecondEmailSent: this.props.siteSelected.SecondEmailReply,
                URL: this.props.siteSelected.URL,
                UserLastModified: this.props.siteSelected.UserLastModified,
                DeletionDate: this.props.siteSelected.DeletionDate,
                RecentViews: this.props.siteSelected.RecentViews
            });
            console.log(this.props.siteSelected["@odata.editLink"]);
        }
    }



    public render() {
        const { firstDayOfWeek } = this.state;
        return (
            <Panel
                isOpen={this.state.showPanel}
                type={PanelType.large}
                onDismiss={this._hidePanel}
                headerText={'Site Deletion Process'}
                closeButtonAriaLabel="Close"
                onRenderFooterContent={this._onRenderFooterContent}
            >
                <Label>{this.state.Title}</Label>
                <Label>{this.state.URL}</Label>
                <Label>{this.state.UserLastModified}</Label>
                <Label>{this.state.RecentViews}</Label>

                <Stack tokens={stackTokens}>

                    <Toggle label="First Email Sent" checked={this.state.FirstEmailSent} onText="Email Sent" offText="Email Unsent" onChange={this._toggleFirstEmailSent} />
                    <Toggle label="First Email Reply" checked={this.state.FirstEmailReply} onText="Email Received" offText="Email not Received" onChange={this._toggleFirstEmailReply} />
                    <Toggle label="Second Email Sent" checked={this.state.SecondEmailSent} onText="Email Sent" offText="Email Unsent" onChange={this._toggleSecondEmailSent} />
                    <Toggle label="Second Email Reply" checked={this.state.SecondEmailReply} onText="Email Received" offText="Email not Received" onChange={this._toggleSecondEmailReply} />
                    <Toggle label="Deletion Confirmed" checked={this.state.Deleted} onText="Site Deleted" offText="Deletion in Process" onChange={this._toggleDeleted} />
                    <Label>Schedule Deletion</Label>
                    <DatePicker initialPickerDate={(this.state.DeletionDate == null || this.state.DeletionDate== "")? new Date(): this._convertSPDate(this.state.DeletionDate) } firstDayOfWeek={firstDayOfWeek} strings={DayPickerStrings} placeholder="Select a date..." ariaLabel="Select a date" style={ {maxWidth: '30vh'}} onSelectDate={this._setDeletionDate}/>
                </Stack>


            </Panel>
        );
    }


    private _onRenderFooterContent = () => {
        return (
            <div>
                <PrimaryButton onClick={
                    () => {
                        this._saveEmail();
                        
                    }
                } style={{ marginRight: '8px' }}>
                    Save Process
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
            Title: "",
            Deleted: null,
            FirstEmailReply: null,
            FirstEmailSent: null,
            SecondEmailReply: null,
            SecondEmailSent: null,
            URL: "",
            UserLastModified: null,
            DeletionDate: "",
            RecentViews: null,
            dateSelected: false
        });
    }
    private _toggleFirstEmailSent = () => {
        if (this.state.FirstEmailSent) {
            this.setState({
                FirstEmailSent: false
            });
        }
        else {
            this.setState({
                FirstEmailSent: true
            });
        }
    }
    private _toggleFirstEmailReply = () => {
        if (this.state.FirstEmailReply) {
            this.setState({
                FirstEmailReply: false
            });
        }
        else {
            this.setState({
                FirstEmailReply: true
            });
        }
    }
    private _toggleSecondEmailSent = () => {
        if (this.state.SecondEmailSent) {
            this.setState({
                SecondEmailSent: false
            });
        }
        else {
            this.setState({
                SecondEmailSent: true
            });
        }
    }
    private _toggleSecondEmailReply = () => {
        if (this.state.SecondEmailReply) {
            this.setState({
                SecondEmailReply: false
            });
        }
        else {
            this.setState({
                SecondEmailReply: true
            });
        }
    }
    private _toggleDeleted = () => {
        if (this.state.Deleted) {
            this.setState({
                Deleted: false
            });
        }
        else {
            this.setState({
                Deleted: true
            });
        }
    }
    private _setDeletionDate = (date: Date | null | undefined): void => {
        this.setState({
            DeletionDate: date.toJSON(),
            dateSelected: true
        });

    }

    private _saveEmail = (): void => {
        var itemType = "SP.Data.TargetSitesForDeletionSteven1234ListItem";

        const body: string = JSON.stringify({
            '__metadata': {
                'type': itemType
            },
            FirstEmailSent: this.state.FirstEmailSent,
            FirstEmailReply: this.state.FirstEmailReply,
            SecondEmailSent: this.state.SecondEmailSent,
            SecondEmailReply: this.state.SecondEmailReply,
            Deleted: this.state.Deleted,
            DeletionDate: (this.state.dateSelected)? this.state.DeletionDate : null
        });
        this.props.spHttpClient.post(`https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/` + this.props.siteSelected["@odata.editLink"], SPHttpClient.configurations.v1, {
            headers: {
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE',
                'Accept': 'application/json;odata=verbose',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''
            },
            body: body
        })
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    console.log("Save response:", response.status);
                }
                this._hidePanel();
                this.props.refreshCallback();
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




}