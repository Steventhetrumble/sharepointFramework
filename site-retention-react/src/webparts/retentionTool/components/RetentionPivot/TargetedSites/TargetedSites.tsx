import * as React from "react";
import {
  ITargetedSitesProps,
  ITargetedSitesState
} from "./TargetedSites.types";
import {
  Stack,
  List,
  Sticky,
  Label,
  StickyPositionType,
  IStackStyles,
  IStackItemStyles,
  ScrollablePane,
  ScrollbarVisibility
} from "office-ui-fabric-react";
import { row } from "../../RetentionTool.styles";
import { TargetedSitesCommandBar } from "./TargetedSitesCommandBar/TargetedSitesCommandBar";
import { ListResult, ListRow } from "../../common/IObjects";
import { TargetedSitesEditPanel } from "./TargetedSitesEditPanel/TargetedSitesEditPanel";
import { TargetedSitesDetailList } from "./TargetedSitesDetailList/TargetedSitesDetailList";
import { TargetedSiteDeletionDialog } from "./TargetedSiteDeletionDialog/TargetedSiteDeletionDialog";

export const wrapper: IStackStyles = {
  root: {
    height: "40vh",
    position: "relative"
  }
};
export const stackItemStyles: IStackItemStyles = {
  root: {
    marginTop: 5,
    marginBottom: 5
  }
};

export class TargetedSites extends React.Component<
  ITargetedSitesProps,
  ITargetedSitesState
> {
  constructor(props: ITargetedSitesProps) {
    super(props);

    this.state = {
      siteSelected: null,
      siteIsSelected: false,
      showPanel: false,
      showDeletionDialog: false
    };
  }
  public render(): JSX.Element {
    return (
      <Stack>
        <Stack>
          <Stack.Item styles={stackItemStyles}>
            <TargetedSitesCommandBar
              onClickEdit={this._onClickEdit}
              onClickDelete={this._onClickDelete}
            ></TargetedSitesCommandBar>
          </Stack.Item>
        </Stack>
        <Stack styles={wrapper}>
          <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
            <Sticky stickyPosition={StickyPositionType.Header}>
              <Stack.Item styles={stackItemStyles}>
                <Label>{"Total Results: " + this.props.totalRows}</Label>
              </Stack.Item>
            </Sticky>
            <Stack.Item>
              <TargetedSitesDetailList
                items={this.props.targets}
                siteSelectedCallback={this._siteSelected}
                siteUnselectedCallback={this._siteUnSelected}
              />
            </Stack.Item>
          </ScrollablePane>
        </Stack>
        <TargetedSitesEditPanel
          siteSelected={this.state.siteSelected}
          onClickSave={this._saveChanges}
          onClickClose={this._hidePanel}
          showPanel={this.state.showPanel}
        ></TargetedSitesEditPanel>
        <TargetedSiteDeletionDialog
          showDeletionDialog={this.state.showDeletionDialog}
          siteSelected={this.state.siteSelected}
          onDeleteConfirmation={this._confirmDeletion}
          onDialogClose={this._closeDialog}
        />
      </Stack>
    );
  }

  private _siteSelected = (site: ListRow): void => {
    this.setState({
      siteSelected: site,
      siteIsSelected: true
    });
  };

  private _siteUnSelected = (): void => {
    this.setState({
      siteSelected: null,
      siteIsSelected: false
    });
  };

  private _onClickEdit = (): void => {
    if (this.state.siteIsSelected) {
      this._showPanel();
      console.log("edit: ", this.state.siteSelected.Title);
    } else {
      alert("you must select a site to Edit.");
    }
  };

  private _onClickDelete = (): void => {
    if (this.state.siteIsSelected) {
      console.log("Delete: ", this.state.siteSelected.Title);
      this._showDialog();
    } else {
      alert("you must select a site to Delete.");
    }
  };

  private _showPanel = (): void => {
    this.setState({
      showPanel: true
    });
  };

  private _hidePanel = (): void => {
    this.setState({
      showPanel: false
    });
  };
  private _showDialog = (): void => {
    this.setState({ showDeletionDialog: true });
  };

  private _closeDialog = (): void => {
    this.setState({ showDeletionDialog: false });
  };

  private _confirmDeletion = (): void => {
    // todo return value from delete site so that can use .then
    this.props.provider
      .deleteSiteFromList(this.state.siteSelected)
      .then((complete: boolean) => {
        if (complete) {
          this.props.onSiteDeletion();
        }
      });
    this._siteUnSelected();
    this._closeDialog();
  };

  private _saveChanges = (modifiedListRow: ListRow): void => {
    this.props.provider.saveTarget(modifiedListRow);
    this._siteUnSelected();
    this._hidePanel();
  };
}
