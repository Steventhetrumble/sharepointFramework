import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { RelevantResults, Table } from '../IManageLists';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { List } from 'office-ui-fabric-react/lib/List';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import {SPTargetedSitesCommandBar} from './SPTargetedSitesCommandBar/SPTargetedSitesCommandBar';
import { oDataQueryNames } from '@microsoft/microsoft-graph-client';
import { string } from 'prop-types';

const classNames = mergeStyleSets({
  


  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px'
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden'
      }
    }
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px'
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px'
  },
  selectionDetails: {
    marginBottom: '20px'
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px'
  }
};

export interface ISPTargetedSitesDetailListState {
  columns: IColumn[];
  items: ListRow[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  announcedMessage?: string;
  selectedSite: ListRow;
  siteIsSelected: boolean;
}

export interface ISPTargetedSitesDetailListProps{
  spHttpClient: SPHttpClient;
    

}
export interface ListResult {
  value: ListRow[];

}

export interface ListRow{
  '@odata.editLink': string;
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
  emails: string;

}

export interface emails {
  
}

export class SPTargetedSitesDetailList extends React.Component<ISPTargetedSitesDetailListProps, ISPTargetedSitesDetailListState> {
  private _selection: Selection;
  private _allitems: ListRow[];
  
  private _getSiteListing(rowStart: number): Promise<ListResult> {
    return this.props.spHttpClient.get( `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetSitesForDeletionSteven1234')/Items?$top=1000`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: ListResult) => {
          console.log(responseJSON.value);
          return responseJSON;
        });
      });
  }

  constructor(props: ISPTargetedSitesDetailListProps) {
    super(props);
    
    this._getSiteListing(0).then((items: ListResult) => {
      this.setState({
        items: items.value
      });
    });
    this._allitems = [];

    const columns: IColumn[] = [
     
      {
        key: 'column1',
        name: 'Title',
        fieldName: 'title',
        minWidth: 60,
        maxWidth: 120,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        onRender:(item: ListRow) => {
          return <span>{item.Title}</span>;
        },
        isPadded: true
      },
      {
        key: 'column2',
        name: 'Date Modified',
        fieldName: 'dateModifiedValue',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: 'number',
        onRender: (item: ListRow) => {
          return <span>{(item.UserLastModified != null) ? this._convertSPDate(item.UserLastModified).toLocaleDateString(): '...'}</span>;
        },
        isPadded: true
      },
      {
        key: 'column3',
        name: 'First Email Sent',
        fieldName: 'modifiedBy',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: ListRow) => {
          return <span>{item.FirstEmailSent ? "yes" : "no"}</span>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'first Email Reply',
        fieldName: 'firstemailreply',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
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
      items: this._allitems,
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      isModalSelection: false,
      isCompactMode: false,
      announcedMessage: undefined,
      selectedSite: null,
      siteIsSelected: false

    };
  }
  private _tempCallBack = (): void => {
    console.log("this level callback"); 
    this._getSiteListing(0).then((items: ListResult) => {
      this._selection.setAllSelected(false);
      this.setState({
        items: items.value,
        selectionDetails: this._getSelectionDetails()
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



  public render() {
    const { columns, isCompactMode, items, selectionDetails, isModalSelection, announcedMessage } = this.state;

    return (
      <Fabric>
        <div className={classNames.controlWrapper}>
          <TextField label="Filter by name:" onChange={this._onChangeText} styles={controlStyles} />
        </div>
        <Stack>
          <Stack.Item>
            <SPTargetedSitesCommandBar 
            callerback={this._tempCallBack}
            spHttpClient={this.props.spHttpClient} 
            siteIsSelected={this.state.siteIsSelected} 
            selectedSite={this.state.selectedSite}></SPTargetedSitesCommandBar>
          
          </Stack.Item>
        
        </Stack>
        {announcedMessage ? <Announced message={announcedMessage} /> : undefined}
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
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            
          />
        </MarqueeSelection>
      </Fabric>
    );
  }

  public componentDidUpdate(previousProps: any, previousState: ISPTargetedSitesDetailListState) {
    if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
      this._selection.setAllSelected(false);
    }
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }


  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    text = text.toLowerCase();
    this.setState({
      items: text ? this._allitems.filter(i => i.Title.toLowerCase().indexOf(text) > -1) : this._allitems
    });
  }


  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        this.setState({
          selectedSite: null,
          siteIsSelected: false
        });
        return 'No items selected';
      case 1:
        this.setState({
          selectedSite: (this._selection.getSelection()[0] as ListRow),
          siteIsSelected: true
        });
        return '1 item selected: ' + (this._selection.getSelection()[0] as ListRow).Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'}`
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}