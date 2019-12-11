import * as React from "react";
import {
  IFindSitesListProps,
  IFindSitesListState
} from "./FindSitesList.types";
import {
  Stack,
  StickyPositionType,
  Sticky,
  Label,
  ScrollablePane,
  ScrollbarVisibility,
  IStackStyles,
  IStackItemStyles,
  IDropdownOption,
  DropdownMenuItemType,
  concatStyleSets
} from "office-ui-fabric-react";
import { FindSitesListCommandBar } from "./FindSitesListCommandBar/FindSitesListCommandBar";
import { FindSitesDetailList } from "./FindSitesDetailList/FindSitesDetailList";
import { SubWebListPanel } from "./SubWebListPanel/SubWebListPanel";
import { GroupChooserModal } from "./GroupChooserModal/GroupChooserModal";
import {
  ISPSite,
  RelevantResults,
  SPGroupList,
  ISPUser
} from "../../common/IObjects";

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

export class FindSitesList extends React.Component<
  IFindSitesListProps,
  IFindSitesListState
> {
  constructor(props: IFindSitesListProps) {
    super(props);

    this.state = {
      showPanel: false,
      showUserModal: false,
      siteIsSelected: false,
      siteSelected: null,
      subWebs: [],
      groups: null,
      allGroups: {},
      displayOptions: [
        {
          key: "fruitsHeader",
          text: "Fruits",
          itemType: DropdownMenuItemType.Header
        },
        { key: "apple", text: "Apple" },
        { key: "banana", text: "Banana" },
        { key: "orange", text: "Orange", disabled: true },
        { key: "grape", text: "Grape" },
        {
          key: "divider_1",
          text: "-",
          itemType: DropdownMenuItemType.Divider
        },
        {
          key: "vegetablesHeader",
          text: "Vegetables",
          itemType: DropdownMenuItemType.Header
        },
        { key: "broccoli", text: "Broccoli" },
        { key: "carrot", text: "Carrot" },
        { key: "lettuce", text: "Lettuce" }
      ],
      idsSelected: [],
      users: []
    };
  }
  public render(): JSX.Element {
    return (
      <Stack>
        <Stack>
          <Stack.Item styles={stackItemStyles}>
            <FindSitesListCommandBar
              onClickInspect={this._showPanel}
            ></FindSitesListCommandBar>
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
              <FindSitesDetailList
                items={this.props.items}
                siteSelectedCallback={this._siteSelected}
                siteUnselectedCallback={this._siteUnSelected}
                alreadyTargetd={this.props.targets}
              />
            </Stack.Item>
          </ScrollablePane>
        </Stack>
        <SubWebListPanel
          subSites={this.state.subWebs}
          onClickAddSite={this._panelAddSite}
          onClickClose={this._hidePanel}
          showPanel={this.state.showPanel}
        ></SubWebListPanel>
        <GroupChooserModal
          groups={this.state.groups}
          showModal={this.state.showUserModal}
          siteSelected={this.state.siteSelected}
          onClickAddTarget={this._modalAddSite}
          onClickCancel={this._modalClickCancel}
          provider={this.props.provider}
          allGroups={this.state.allGroups}
          displayOptions={this.state.displayOptions}
          onChangeSelectedIds={this._onSelectedIdsChange}
          selectedIds={this.state.idsSelected}
          users={this.state.users}
        />
      </Stack>
    );
  }

  private _siteSelected = (site: ISPSite): void => {
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

  private _panelAddSite = (): void => {
    this._lookupGroupsForModal();
    this._hidePanel();
    this._showModal();
  };

  private _modalAddSite = (): void => {
    if (this.state.users.length > 0) {
      this._hideModal();
      this.props.provider.postListData(
        this.state.siteSelected,
        this.state.users
      );
      this.props.onSiteAddition();
    } else {
      alert("bro");
    }
  };

  private _modalClickCancel = (): void => {
    this._hideModal();
    this._showPanel();
  };

  private _showModal = (): void => {
    this.setState({
      showUserModal: true
    });
  };

  private _hideModal = (): void => {
    this.setState({
      showUserModal: false,
      groups: null,
      allGroups: {},
      displayOptions: [
        {
          key: "fruitsHeader",
          text: "Fruits",
          itemType: DropdownMenuItemType.Header
        },
        { key: "apple", text: "Apple" },
        { key: "banana", text: "Banana" },
        { key: "orange", text: "Orange", disabled: true },
        { key: "grape", text: "Grape" },
        {
          key: "divider_1",
          text: "-",
          itemType: DropdownMenuItemType.Divider
        },
        {
          key: "vegetablesHeader",
          text: "Vegetables",
          itemType: DropdownMenuItemType.Header
        },
        { key: "broccoli", text: "Broccoli" },
        { key: "carrot", text: "Carrot" },
        { key: "lettuce", text: "Lettuce" }
      ],
      idsSelected: [],
      users: []
    });
  };

  private _showPanel = (): void => {
    if (this.state.siteIsSelected) {
      this.setState({
        showPanel: true
      });
      this._lookupSubWebsForPanel();
      console.log("edit: ", this.state.siteSelected.Title);
    } else {
      alert("you must select a site to inspect.");
    }
  };

  private _hidePanel = (): void => {
    this.setState({
      showPanel: false,
      subWebs: []
    });
  };

  private _lookupGroupsForModal = (): void => {
    this.props.provider
      .getUserGroups(this.state.siteSelected)
      .then((_groups: SPGroupList) => {
        this.setState({
          groups: _groups
        });
        this._createDropdownOptions();
      });
  };

  private _lookupSubWebsForPanel = (): void => {
    this.props.provider
      .getSubWebs(this.state.siteSelected)
      .then((res: RelevantResults) => {
        this.props.provider
          .lookupDates(res)
          .forEach((prom: Promise<ISPSite>) => {
            prom.then((site: ISPSite) => {
              let tempItems: ISPSite[] = this.state.subWebs;
              this.setState({
                subWebs: [...tempItems, site]
              });
            });
          });
      });
  };

  private _createDropdownOptions(): void {
    this.props.provider
      .makeOptions(this.state.groups, this.state.siteSelected.Url)
      .then((results: [IDropdownOption[], {}]) => {
        console.log("this is working");
        this.setState({
          displayOptions: results[0],
          allGroups: results[1]
        });
      })
      .catch(err => {
        console.log("this isnt working");
      });
  }

  private _onSelectedIdsChange = (ids: number[]): void => {
    this.setState({
      idsSelected: ids
    });
    this._getUsersFromGroups(ids);
  };

  private _getUsersFromGroups(id: number[]): void {
    if (id[0] != null) {
      console.log(id);
      console.log(this.state.allGroups);

      var tempItems: ISPUser[] = [];

      id.forEach((num: number) => {
        this.state.allGroups[num].value.forEach((user: ISPUser) => {
          if (user.Email !== "") {
            tempItems.push(user);
          }
        });
      });
      this.setState({
        users: tempItems
      });
    }
  }
}
