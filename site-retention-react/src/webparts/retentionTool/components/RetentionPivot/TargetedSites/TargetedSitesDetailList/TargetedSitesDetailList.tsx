import * as React from "react";

import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  ConstrainMode,
  IDetailsHeaderProps
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import { IList, ListRow, ISPDate, ListResult } from "../../../common/IObjects";
import {
  ITargetedSitesDetailListProps,
  ITargetedSitesDetailListState,
  TargetedSites
} from "./TargetedSitesDetailList.types";
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
import { Label } from "office-ui-fabric-react";

const classNames = mergeStyleSets({
  wrapper: {
    position: "relative"
  }
});
const controlStyles = {
  root: {
    color: "red",
    fontWeight: "bold"
  }
};
const controlStyles2 = {
  root: {
    color: "primary"
  }
};

export class TargetedSitesDetailList extends React.Component<
  ITargetedSitesDetailListProps,
  ITargetedSitesDetailListState
> {
  private _selection: Selection;

  constructor(props: ITargetedSitesDetailListProps) {
    super(props);

    if (Environment.type === EnvironmentType.Local) {
    } else {
      const columns: IColumn[] = [
        {
          key: "column1",
          name: "Title",
          fieldName: "title",
          minWidth: 60,
          maxWidth: 120,
          isRowHeader: true,
          isResizable: true,
          isSorted: true,
          isSortedDescending: false,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
          onColumnClick: this._onColumnClick,
          data: "string",
          onRender: (item: ListRow) => {
            return <span>{item.Title}</span>;
          },
          isPadded: true
        },
        {
          key: "column2",
          name: "Date Modified",
          fieldName: "dateModifiedValue",
          minWidth: 70,
          maxWidth: 90,
          isResizable: true,
          onColumnClick: this._onColumnClick,
          data: "number",
          onRender: (item: ListRow) => {
            //   this._convertSPDate(item.UserLastModified).toLocaleDateString()
            return (
              <span>
                {item.UserLastModified != null ? item.UserLastModified : "..."}
              </span>
            );
          },
          isPadded: true
        },
        {
          key: "column3",
          name: "First Email Sent",
          fieldName: "modifiedBy",
          minWidth: 70,
          maxWidth: 90,
          isResizable: true,
          isCollapsible: true,
          data: "string",
          onColumnClick: this._onColumnClick,
          onRender: (item: ListRow) => {
            return <span>{item.FirstEmailSent ? "yes" : "no"}</span>;
          },
          isPadded: true
        },
        {
          key: "column4",
          name: "first Email Reply",
          fieldName: "firstemailreply",
          minWidth: 70,
          maxWidth: 90,
          isResizable: true,
          isCollapsible: true,
          data: "string",
          onColumnClick: this._onColumnClick,
          onRender: (item: ListRow) => {
            return <span>{item.FirstEmailReply ? "yes" : "no"}</span>;
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
    const { columns, items, selectionDetails, announcedMessage } = this.state;

    return (
      <div className={classNames.wrapper}>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={this.props.items.value}
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
          targetedSite: this._selection.getSelection()[0] as ListRow,
          targetSiteSelected: true
        });
        this.props.siteSelectedCallback(
          this._selection.getSelection()[0] as ListRow
        );
        return (
          "1 item selected: " +
          (this._selection.getSelection()[0] as ListRow).Title
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
    const newValues = this._copyAndSort(
      items.value,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    let newItems: ListResult = {
      value: newValues
    };

    this.setState({
      columns: newColumns,
      items: newItems
    });
  };
  public componentDidUpdate(previousProps: ITargetedSitesDetailListProps) {
    if (previousProps.items != this.props.items) {
      this.setState({
        items: null
      });
      this.setState({
        items: this.props.items
      });
      this._selection.setAllSelected(false);
    }
  }

  private _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    const key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
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
