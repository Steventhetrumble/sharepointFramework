import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, ConstrainMode, IDetailsHeaderProps, } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import {
  IList,
  ISPSite,
  ISPDate
} from '../../common/IObjects';
import {
  Result,
  RelevantResults,
  Row
} from '../IManageLists';

import { IDataProvider } from '../../dataproviders/IDataProvider';
import {
  Version,
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import styles from '../ManageLists.module.scss';
import { SPSearchPanel } from './SPSearchPanel/SPSearchPanel';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from 'office-ui-fabric-react/lib/Sticky';
import { IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { TooltipHost, ITooltipHostProps } from 'office-ui-fabric-react/lib/Tooltip';
import { Label } from 'office-ui-fabric-react';

const classNames = mergeStyleSets({
  wrapper: {
    height: '80vh',
    position: 'relative'
  },


  selectionDetails: {
    marginLeft: '5px',
    marginBottom: '20px'
  }
});
const controlStyles = {
  root: {
    color: 'red',
    fontWeight: 'bold'
  }
};
const controlStyles2 = {
  root: {
   color: 'primary'
  }
};

export interface ListRows {
  value: ListRow[];
}

export interface ListRow {
  Title: string;
}

export interface ISPDetailListState {
  columns: IColumn[];
  items: ISPSite[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  announcedMessage?: string;
  lists: IList[];
  rootSite: ISPSite;
  totalRows: number;
  targetedSite: ISPSite;
  targetSiteSelected: boolean;
  sitesAlreadyTargeted: TargetedSites;
}

export interface ISPDetailListProps {
  provider: IDataProvider;
  site: ISPSite;
  spHttpClient: SPHttpClient;
  
}

export interface TargetedSites {
  [key: string]: boolean;
}


export class SPDetailList extends React.Component<ISPDetailListProps, ISPDetailListState> {
  private _selection: Selection;

  constructor(props: ISPDetailListProps) {
    super(props);

    if (Environment.type === EnvironmentType.Local) {
      this.props.provider.getAllLists().then((_lists: IList[]) => {


        console.log(this.state.lists);
      });
    }
    else {
      this._getTargetedSites().then((value: boolean) => {
        

      });
      this._getSiteListing(0).then((result: RelevantResults) => {
        let _sites: ISPSite[] = [];

        const randomDate = _randomDate(new Date(2012, 0, 1), new Date(2012, 0, 1));

        result.Table.Rows.forEach((row: Row) => {
          if(this.state.sitesAlreadyTargeted[row.Cells[2].Value]){
            console.log(row.Cells[2].Value);
          }
          this._getSPDate(row).then((responseDate: ISPDate) => {
           
            _sites.push({
              DocID: row.Cells[1].Value,
              Title: row.Cells[2].Value,
              Url: row.Cells[3].Value,
              ViewsLifeTime: Number(row.Cells[6].Value),
              ViewsRecent: Number(row.Cells[7].Value),
              Size: Number(row.Cells[8].Value),
              SiteDescription: row.Cells[9].Value,
              LastItemUserModifiedDateSharepoint: (responseDate.error == null) ? responseDate.value : "ok",
              LastItemUserModifiedDate: (responseDate.error == null) ? responseDate.date : randomDate.date,
              LastItemUserModifiedDatevalue: (responseDate.error == null) ? responseDate.datevalue : randomDate.value,
              LastItemUserModifiedDateFomatted: (responseDate.error == null) ? responseDate.dateFormatted : randomDate.dateFormatted,
              renderTemplateId: row.Cells[16].Value
            });
          });


        });
        for (var i = 0; i < Math.floor(result.TotalRows / 500); i++) {
          this._getSiteListing(500 + i * 500).then((result2: RelevantResults) => {
            result2.Table.Rows.forEach((row: Row) => {
              if(this.state.sitesAlreadyTargeted[row.Cells[2].Value]){
                console.log(row.Cells[2].Value);
              }
              this._getSPDate(row).then((responseDate: ISPDate) => {
                _sites.push({
                  DocID: row.Cells[1].Value,
                  Title: row.Cells[2].Value,
                  Url: row.Cells[3].Value,
                  ViewsLifeTime: Number(row.Cells[6].Value),
                  ViewsRecent: Number(row.Cells[7].Value),
                  Size: Number(row.Cells[8].Value),
                  SiteDescription: row.Cells[9].Value,
                  LastItemUserModifiedDateSharepoint: (responseDate.error == null) ? responseDate.value : "ok",
                  LastItemUserModifiedDate: (responseDate.error == null) ? responseDate.date : randomDate.date,
                  LastItemUserModifiedDatevalue: (responseDate.error == null) ? responseDate.datevalue : randomDate.value,
                  LastItemUserModifiedDateFomatted: (responseDate.error == null) ? responseDate.dateFormatted : randomDate.dateFormatted,
                  renderTemplateId: row.Cells[16].Value

                });
              });
            });
          });
        }

        this.setState({
          items: _sites
        });
      });


    }

    const columns: IColumn[] = [

      {
        key: 'column1',
        name: 'Title',
        fieldName: 'title',
        minWidth: 55,
        maxWidth: 140,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        onRender: (item: ISPSite) => {
          return <Label styles={(this.state.sitesAlreadyTargeted[item.Title])?controlStyles: controlStyles2}>{item.Title}</Label>;
        },
        isPadded: true
      },
      {
        key: 'column2',
        name: 'Url',
        fieldName: 'url',
        minWidth: 105,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        onRender: (item: ISPSite) => {
          return <Label styles={(this.state.sitesAlreadyTargeted[item.Title])?controlStyles: controlStyles2} >{item.Url}</Label>;
        },
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Last Item User Mod',
        fieldName: 'lastitemusermodified',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: 'number',
        onRender: (item: ISPSite) => {
          // TODO will need to query site for this info
          return <Label styles={(this.state.sitesAlreadyTargeted[item.Title])?controlStyles: controlStyles2}>{item.LastItemUserModifiedDateFomatted}</Label>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Views Lifetime',
        fieldName: 'viewslifetime',
        minWidth: 20,
        maxWidth: 40,
        isResizable: true,
        isCollapsible: true,
        data: 'number',
        onColumnClick: this._onColumnClick,
        onRender: (item: ISPSite) => {
          return <Label styles={(this.state.sitesAlreadyTargeted[item.Title])?controlStyles: controlStyles2}>{item.ViewsLifeTime}</Label>;
        },
        isPadded: true
      },
      {
        key: 'column5',
        name: 'Views Recent',
        fieldName: 'viewsrecent',
        minWidth: 20,
        maxWidth: 40,
        isResizable: true,
        isCollapsible: true,
        data: 'number',
        onColumnClick: this._onColumnClick,
        onRender: (item: ISPSite) => {
          return <Label styles={(this.state.sitesAlreadyTargeted[item.Title])?controlStyles: controlStyles2}>{item.ViewsRecent}</Label>;
        },
        isPadded: true
      },
      {
        key: 'column6',
        name: 'File Size',
        fieldName: 'fileSizeRaw',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'number',
        onColumnClick: this._onColumnClick,
        
        onRender: (item: ISPSite) => {
          return <Label styles={(this.state.sitesAlreadyTargeted[item.Title])?controlStyles: controlStyles2}>{item.Size}</Label>;
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
      items: [],
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      isModalSelection: false,
      isCompactMode: false,
      announcedMessage: undefined,
      lists: [],
      rootSite: this.props.site,
      totalRows: 0,
      targetedSite: null,
      targetSiteSelected: false,
      sitesAlreadyTargeted: {}
    };
  }



  public render() {
    const { columns, isCompactMode, items, selectionDetails, isModalSelection, announcedMessage } = this.state;

    return (

      <Fabric>
        <div className={classNames.wrapper}>
          <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
            <Sticky stickyPosition={StickyPositionType.Header}>
              <div className={styles.row}>Total Results: {this.state.totalRows}</div>
              <div className={classNames.selectionDetails}>{selectionDetails}</div>
              {announcedMessage ? <Announced message={announcedMessage} /> : undefined}
              <SPSearchPanel 
              spHttpClient={this.props.spHttpClient} 
              showPanel={this.state.targetSiteSelected} 
              targetedSite={this.state.targetedSite}
              sitesAlreadyTargeted={this.state.sitesAlreadyTargeted} 
              siteAddedCallback={this._siteAddedCallback}
              />
            </Sticky>
            <MarqueeSelection selection={this._selection}>

              <DetailsList
                items={items}
                compact={isCompactMode}
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
          </ScrollablePane>


        </div>

      </Fabric>
    );
  }

  private _getTargetedSites(): Promise<boolean> {
    return this.props.spHttpClient.get(`https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetSitesForDeletionSteven1234')/Items?$top=1000`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responesJSON: ListRows) => {
          let _targetedSites: TargetedSites = {};
          responesJSON.value.forEach((row: ListRow) => {
            _targetedSites[row.Title] = true;
          });
          this.setState({
            sitesAlreadyTargeted: _targetedSites
          });
          return true;
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

  private _getSiteListing(rowStart: number): Promise<RelevantResults> {
    return this.props.spHttpClient.get(this.props.site.Url + `/_api/search/query?querytext=%27(contentclass:STS_Site) Path:https://rcirogers.sharepoint.com/sites/* OR Path:https://rcirogers.sharepoint.com/teams/* NOT (WebTemplate:GROUP) %27&trimduplicates=false&RowLimit=500&startrow=${rowStart.toString()}&selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate,WebTemplate%27`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: Result) => {
          console.log(responseJSON);
          this.setState({
            totalRows: responseJSON.PrimaryQueryResult.RelevantResults.TotalRows
          });
          return responseJSON.PrimaryQueryResult.RelevantResults;
        });
      });
  }

  public componentDidUpdate(previousProps: any, previousState: ISPDetailListState) {
    if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
      this._selection.setAllSelected(false);
    }
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
        return 'No items selected';
      case 1:
        this.setState({
          targetedSite: (this._selection.getSelection()[0] as ISPSite),
          targetSiteSelected: true
        });
        return '1 item selected: ' + (this._selection.getSelection()[0] as ISPSite).Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const columns = this.state.columns;
    const items = this.state.items;
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
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
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

  private _copyAndSort(items: ISPSite[], columnKey: string, isSortedDescending?: boolean) {
    this._selection.setAllSelected(false);
    console.log(columnKey);
    const key = columnKey as keyof ISPSite;
    if (columnKey === "lastitemusermodified") {
      return items.sort((a: ISPSite, b: ISPSite) => (
        (isSortedDescending ? a.LastItemUserModifiedDatevalue > b.LastItemUserModifiedDatevalue : b.LastItemUserModifiedDatevalue > a.LastItemUserModifiedDatevalue) ? 1 : -1)
      );
    }
    else if (columnKey === "title") {
      return items.sort((a: ISPSite, b: ISPSite) => (
        (isSortedDescending ? a.Title.toLowerCase() > b.Title.toLowerCase() : a.Title.toLowerCase() < b.Title.toLowerCase()) ? 1 : -1)
      );
    }
    else if (columnKey === "url") {
      return items.sort((a: ISPSite, b: ISPSite) => (
        (isSortedDescending ? a.Url.toLowerCase() > b.Url.toLowerCase() : a.Url.toLowerCase() < b.Url.toLowerCase()) ? 1 : -1)
      );
    }
    else if (columnKey === "viewslifetime") {
      return items.sort((a: ISPSite, b: ISPSite) => (
        (isSortedDescending ? a.ViewsLifeTime > b.ViewsLifeTime : b.ViewsLifeTime > a.ViewsLifeTime) ? 1 : -1)
      );
    }
    else if (columnKey === "viewsrecent") {
      return items.sort((a: ISPSite, b: ISPSite) => (
        (isSortedDescending ? a.ViewsRecent > b.ViewsRecent : a.ViewsRecent < b.ViewsRecent) ? 1 : -1)
      );
    }
    else if (columnKey === "filesSizeRaw") {
      return items.sort((a: ISPSite, b: ISPSite) => (
        (isSortedDescending ? a.Size > b.Size : a.Size < b.Size) ? 1 : -1)
      );
    }
  
    else {
      return items.slice(0).sort((a: ISPSite, b: ISPSite) => (
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1)
      );
    }
  
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

function onRenderDetailsHeader(props: IDetailsHeaderProps, defaultRender?: IRenderFunction<IDetailsHeaderProps>): JSX.Element {
  return (
    <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
      {defaultRender!({
        ...props,
        onRenderColumnHeaderTooltip: (tooltipHostProps: ITooltipHostProps) => <TooltipHost {...tooltipHostProps} />
      })}
    </Sticky>
  );
}