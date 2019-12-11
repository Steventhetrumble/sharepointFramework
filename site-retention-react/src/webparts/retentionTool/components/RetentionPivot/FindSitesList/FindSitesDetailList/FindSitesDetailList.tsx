import * as React from "react";

import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  ConstrainMode,
  IDetailsHeaderProps,
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { mergeStyleSets, DefaultPalette } from "office-ui-fabric-react/lib/Styling";
import { IList, ISPSite, ISPDate } from "../../../common/IObjects";
import {IFindSitesDetailListProps, IFindSitesDetailListState, TargetedSites} from "./FindSitesDetailList.types"
import {
  Version,
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";
import { SPHttpClientResponse, SPHttpClient } from "@microsoft/sp-http";
import { Sticky, StickyPositionType } from "office-ui-fabric-react/lib/Sticky";
import { IRenderFunction } from "office-ui-fabric-react/lib/Utilities";
import {
  TooltipHost,
  ITooltipHostProps
} from "office-ui-fabric-react/lib/Tooltip";
import { Label, Text } from "office-ui-fabric-react";

const classNames = mergeStyleSets({
  wrapper: {
    height: "80vh",
    position: "relative"
  },
  
  
});
const controlStyles = {
  root: {
    color: "red",
    fontWeight: "bold",
    verticalAlign: "start"
  }
};
const controlStyles2 = {
  root: {
    color: "primary",
    verticalAlign: "start"
  }
};

const checkbox = {
  root: {
    verticalAlign: "center"
  }
};



export class FindSitesDetailList extends React.Component<
  IFindSitesDetailListProps,
  IFindSitesDetailListState
> {
  private _selection: Selection;

  constructor(props: IFindSitesDetailListProps) {
    super(props);

    if (Environment.type === EnvironmentType.Local) {
    } else {
      const columns: IColumn[] = [
        {
          key: "column1",
          name: "Title",
          fieldName: "title",
          minWidth: 55,
          maxWidth: 140,
          isRowHeader: true,
          isResizable: true,
          isSorted: true,
          isSortedDescending: false,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
          onColumnClick: this._onColumnClick,
          data: "string",
          onRender: (item: ISPSite) => {
            return (
              <Text
                styles={
                  this.state.sitesAlreadyTargeted[item.Title]
                    ? controlStyles
                    : controlStyles2
                }
              >
                {item.Title}
              </Text>
            );
          },
          isPadded: true
        },
        {
          key: "column2",
          name: "Url",
          fieldName: "url",
          minWidth: 105,
          maxWidth: 150,
          isRowHeader: true,
          isResizable: true,
          isSorted: true,
          isSortedDescending: false,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
          onColumnClick: this._onColumnClick,
          data: "string",
          onRender: (item: ISPSite) => {
            return (
              <Text
                styles={
                  this.state.sitesAlreadyTargeted[item.Title]
                    ? controlStyles
                    : controlStyles2
                }
              >
                {item.Url}
              </Text>
            );
          },
          isPadded: true
        },
        {
          key: "column3",
          name: "Last Item User Mod",
          fieldName: "lastitemusermodified",
          minWidth: 70,
          maxWidth: 90,
          isResizable: true,
          onColumnClick: this._onColumnClick,
          data: "number",
          onRender: (item: ISPSite) => {
            // TODO will need to query site for this info
            return (
              <Text
                styles={
                  this.state.sitesAlreadyTargeted[item.Title]
                    ? controlStyles
                    : controlStyles2
                }
              >
                {item.LastItemUserModifiedDateFomatted}
              </Text>
            );
          },
          isPadded: true
        },
        {
          key: "column4",
          name: "Views Lifetime",
          fieldName: "viewslifetime",
          minWidth: 20,
          maxWidth: 40,
          isResizable: true,
          isCollapsible: true,
          data: "number",
          onColumnClick: this._onColumnClick,
          onRender: (item: ISPSite) => {
            return (
              <Text
                styles={
                  this.state.sitesAlreadyTargeted[item.Title]
                    ? controlStyles
                    : controlStyles2
                }
              >
                {item.ViewsLifeTime}
              </Text>
            );
          },
          isPadded: true
        },
        {
          key: "column5",
          name: "Views Recent",
          fieldName: "viewsrecent",
          minWidth: 20,
          maxWidth: 40,
          isResizable: true,
          isCollapsible: true,
          data: "number",
          onColumnClick: this._onColumnClick,
          onRender: (item: ISPSite) => {
            return (
              <Text
                styles={
                  this.state.sitesAlreadyTargeted[item.Title]
                    ? controlStyles
                    : controlStyles2
                }
              >
                {item.ViewsRecent}
              </Text>
            );
          },
          isPadded: true
        },
        {
          key: "column6",
          name: "File Size",
          fieldName: "fileSizeRaw",
          minWidth: 70,
          maxWidth: 90,
          isResizable: true,
          isCollapsible: true,
          data: "number",
          onColumnClick: this._onColumnClick,

          onRender: (item: ISPSite) => {
            return (
              <Text
                styles={
                  this.state.sitesAlreadyTargeted[item.Title]
                    ? controlStyles
                    : controlStyles2
                }
              >
                {item.Size}
              </Text>
            );
          }
        }
      ];

      this._selection = new Selection({
        onSelectionChanged: () => {
          this.setState({
            selectionDetails: this._getSelectionDetails()
          });
        }
      });

      this.state = {
        items: this.props.items,
        columns: columns,
        selectionDetails: this._getSelectionDetails(),
        lists: [],
        targetedSite: null,
        targetSiteSelected: false,
        sitesAlreadyTargeted: {}
      };
    }
  }

  
  public render() {
    const {
      columns,
      items,
      selectionDetails,
      announcedMessage
    } = this.state;

    return (
        <div className={classNames.wrapper}> 
            <MarqueeSelection selection={this._selection}>
              <DetailsList
                items={items}
                compact={true}
                columns={columns}
                selectionMode={SelectionMode.single}
                getKey={this._getKey}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                selection={this._selection}
                selectionPreservedOnEmptyClick={true}
                constrainMode={ConstrainMode.unconstrained}
                onRenderDetailsHeader={onRenderDetailsHeader}
                onItemInvoked={this._onItemInvoked}
                enterModalSelectionOnTouch={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="Row checkbox"
              />
            </MarqueeSelection>
        </div>
    );
  }


  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.Title}`);
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        this.setState({
          targetedSite: null,
          targetSiteSelected: false
        });
        this.props.siteUnselectedCallback();
        return "No items selected";
      case 1:
        this.setState({
          targetedSite: this._selection.getSelection()[0] as ISPSite,
          targetSiteSelected: true
        });
        this.props.siteSelectedCallback(this._selection.getSelection()[0] as ISPSite);
        return (
          "1 item selected: " +
          (this._selection.getSelection()[0] as ISPSite).Title
        );
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const columns = this.state.columns;
    const items = this.state.items;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(
      currCol => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${
            currColumn.isSortedDescending ? "descending" : "ascending"
          }`
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(
      items,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }
  public componentDidUpdate(previousProps: IFindSitesDetailListProps){
      if(previousProps.items != this.props.items){
        this.setState({
            items: this.props.items
        });
      }
  }


  private _siteAddedCallback = (s: string): void => {
    this._selection.setAllSelected(false);
    let tempSites: TargetedSites = this.state.sitesAlreadyTargeted;
    tempSites[s] = true;
    this.setState({
      sitesAlreadyTargeted: tempSites
    });
    this.setState({
      items: this.state.items
    });
  }

  private _copyAndSort(
    items: ISPSite[],
    columnKey: string,
    isSortedDescending?: boolean
  ) {
    this._selection.setAllSelected(false);
    console.log(columnKey);
    const key = columnKey as keyof ISPSite;
    if (columnKey === "lastitemusermodified") {
      return items.sort((a: ISPSite, b: ISPSite) =>
        (isSortedDescending
        ? a.LastItemUserModifiedDatevalue > b.LastItemUserModifiedDatevalue
        : b.LastItemUserModifiedDatevalue > a.LastItemUserModifiedDatevalue)
          ? 1
          : -1
      );
    } else if (columnKey === "title") {
      return items.sort((a: ISPSite, b: ISPSite) =>
        (isSortedDescending
        ? a.Title.toLowerCase() > b.Title.toLowerCase()
        : a.Title.toLowerCase() < b.Title.toLowerCase())
          ? 1
          : -1
      );
    } else if (columnKey === "url") {
      return items.sort((a: ISPSite, b: ISPSite) =>
        (isSortedDescending
        ? a.Url.toLowerCase() > b.Url.toLowerCase()
        : a.Url.toLowerCase() < b.Url.toLowerCase())
          ? 1
          : -1
      );
    } else if (columnKey === "viewslifetime") {
      return items.sort((a: ISPSite, b: ISPSite) =>
        (isSortedDescending
        ? a.ViewsLifeTime > b.ViewsLifeTime
        : b.ViewsLifeTime > a.ViewsLifeTime)
          ? 1
          : -1
      );
    } else if (columnKey === "viewsrecent") {
      return items.sort((a: ISPSite, b: ISPSite) =>
        (isSortedDescending
        ? a.ViewsRecent > b.ViewsRecent
        : a.ViewsRecent < b.ViewsRecent)
          ? 1
          : -1
      );
    } else if (columnKey === "filesSizeRaw") {
      return items.sort((a: ISPSite, b: ISPSite) =>
        (isSortedDescending ? a.Size > b.Size : a.Size < b.Size) ? 1 : -1
      );
    } else {
      return items
        .slice(0)
        .sort((a: ISPSite, b: ISPSite) =>
          (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
        );
    }
  }
 
}



function onRenderDetailsHeader(
  props: IDetailsHeaderProps,
  defaultRender?: IRenderFunction<IDetailsHeaderProps>
): JSX.Element {
  return (
    <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
      {defaultRender!({
        ...props,
        onRenderColumnHeaderTooltip: (tooltipHostProps: ITooltipHostProps) => (
          <TooltipHost {...tooltipHostProps} />
        )
      })}
    </Sticky>
  );
}
