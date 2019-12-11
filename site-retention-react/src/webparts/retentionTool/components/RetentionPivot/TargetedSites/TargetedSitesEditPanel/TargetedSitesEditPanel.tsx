import * as React from "react";
import {
  DefaultButton,
  PrimaryButton
} from "office-ui-fabric-react/lib/Button";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import {
  ITargetedSitesEditPanelProps,
  ITargetedSitesEditPanelState,
  DayPickerStrings
} from "./TargetedSitesEditPanel.types";
import {
  IDatePickerStrings,
  DayOfWeek,
  Label,
  Stack,
  Toggle,
  DatePicker,
  IStackTokens,
  Text
} from "office-ui-fabric-react";
import {
  title,
  stackStyles,
  itemAlignmentsStackTokens,
  stackItemStyles
} from "./TargetedSitesEditPanel.styles";
import { ListRow } from "../../../common/IObjects";

const stackTokens: IStackTokens = { childrenGap: 10 };

export class TargetedSitesEditPanel extends React.Component<
  ITargetedSitesEditPanelProps,
  ITargetedSitesEditPanelState
> {
  constructor(props: ITargetedSitesEditPanelProps) {
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

  public componentDidUpdate(previousProps: ITargetedSitesEditPanelProps) {
    if (
      previousProps.siteSelected != this.props.siteSelected &&
      this.props.siteSelected != null
    ) {
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
    }
  }

  public render() {
    return (
      <div>
        <Panel
          isOpen={this.props.showPanel}
          closeButtonAriaLabel="Close"
          onDismiss={this.props.onClickClose}
          type={PanelType.large}
          headerText="Large Panel"
          onRenderFooterContent={this._onRenderFooterContent}
          onRenderHeader={this._onRenderHeader}
        >
          <Stack horizontal wrap styles={stackStyles} tokens={itemAlignmentsStackTokens}>
            <Stack.Item styles={stackItemStyles}>
              <Label>{"Target Name: " + this.state.Title}</Label>
            </Stack.Item>
            <Stack.Item styles={stackItemStyles}>
              <Label>{"Url: " + this.state.URL}</Label>
            </Stack.Item>
            <Stack.Item styles={stackItemStyles}>
              <Label>
                {"User Last Modified: " + this.state.UserLastModified}
              </Label>
            </Stack.Item>
            <Stack.Item styles={stackItemStyles}>
              <Label>{"Recent Views: " + this.state.RecentViews}</Label>
            </Stack.Item>
          </Stack>

          <Stack tokens={stackTokens}>
            <Toggle
              label="First Email Sent"
              checked={this.state.FirstEmailSent}
              onText="Email Sent"
              offText="Email Unsent"
              onChange={this._toggleFirstEmailSent}
            />
            <Toggle
              label="First Email Reply"
              checked={this.state.FirstEmailReply}
              onText="Email Received"
              offText="Email not Received"
              onChange={this._toggleFirstEmailReply}
            />
            <Toggle
              label="Second Email Sent"
              checked={this.state.SecondEmailSent}
              onText="Email Sent"
              offText="Email Unsent"
              onChange={this._toggleSecondEmailSent}
            />
            <Toggle
              label="Second Email Reply"
              checked={this.state.SecondEmailReply}
              onText="Email Received"
              offText="Email not Received"
              onChange={this._toggleSecondEmailReply}
            />
            <Toggle
              label="Deletion Confirmed"
              checked={this.state.Deleted}
              onText="Site Deleted"
              offText="Deletion in Process"
              onChange={this._toggleDeleted}
            />
            <Label>Schedule Deletion</Label>
            <DatePicker
              initialPickerDate={
                this.state.DeletionDate == null || this.state.DeletionDate == ""
                  ? new Date()
                  : this._convertSPDate(this.state.DeletionDate)
              }
              firstDayOfWeek={this.state.firstDayOfWeek}
              strings={DayPickerStrings}
              placeholder="Select a date..."
              ariaLabel="Select a date"
              style={{ maxWidth: "30vh" }}
              onSelectDate={this._setDeletionDate}
            />
          </Stack>
        </Panel>
      </div>
    );
  }

  private _onRenderFooterContent = () => {
    return (
      <div>
        <PrimaryButton
          onClick={() => {
            this._save();
          }}
          style={{ marginRight: "8px" }}
        >
          Save Changes
        </PrimaryButton>
        <DefaultButton onClick={this.props.onClickClose}>Cancel</DefaultButton>
      </div>
    );
  };
  private _onRenderHeader = () => {
    return (
      <Stack styles={stackStyles} tokens={itemAlignmentsStackTokens}>
        <Stack.Item align="start" styles={stackItemStyles}>
          <Text styles={title}>Deletion Process</Text>
        </Stack.Item>
      </Stack>
    );
  };

  private _save = () => {
    let modifiedLR: ListRow = {
      "@odata.editLink": this.props.siteSelected["@odata.editLink"],
      Title: this.props.siteSelected.Title,
      Deleted: this.state.Deleted,
      FirstEmailReply: this.state.FirstEmailReply,
      FirstEmailSent: this.state.FirstEmailSent,
      SecondEmailReply: this.state.SecondEmailReply,
      SecondEmailSent: this.state.SecondEmailSent,
      URL: this.props.siteSelected.URL,
      UserLastModified: this.props.siteSelected.UserLastModified,
      DeletionDate: this.state.DeletionDate,
      RecentViews: this.props.siteSelected.RecentViews,
      emails: this.props.siteSelected.emails
    };
    this.props.onClickSave(modifiedLR);
  };

  private _toggleFirstEmailSent = () => {
    if (this.state.FirstEmailSent) {
      this.setState({
        FirstEmailSent: false
      });
    } else {
      this.setState({
        FirstEmailSent: true
      });
    }
  };
  private _toggleFirstEmailReply = () => {
    if (this.state.FirstEmailReply) {
      this.setState({
        FirstEmailReply: false
      });
    } else {
      this.setState({
        FirstEmailReply: true
      });
    }
  };
  private _toggleSecondEmailSent = () => {
    if (this.state.SecondEmailSent) {
      this.setState({
        SecondEmailSent: false
      });
    } else {
      this.setState({
        SecondEmailSent: true
      });
    }
  };
  private _toggleSecondEmailReply = () => {
    if (this.state.SecondEmailReply) {
      this.setState({
        SecondEmailReply: false
      });
    } else {
      this.setState({
        SecondEmailReply: true
      });
    }
  };
  private _toggleDeleted = () => {
    if (this.state.Deleted) {
      this.setState({
        Deleted: false
      });
    } else {
      this.setState({
        Deleted: true
      });
    }
  };

  private _setDeletionDate = (date: Date | null | undefined): void => {
    this.setState({
      DeletionDate: date.toJSON(),
      dateSelected: true
    });
  };

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
    let dDate: Date = new Date(
      Number(xYear),
      Number(xMonth) - 1,
      Number(xDay),
      Number(xHour),
      Number(xMin),
      Number(xSec)
    );
    return dDate;
  }
}
