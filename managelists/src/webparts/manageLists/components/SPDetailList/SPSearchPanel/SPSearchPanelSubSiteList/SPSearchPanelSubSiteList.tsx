import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { ISPSite } from '../../../../common/IObjects';
import { Modal, IDragOptions } from 'office-ui-fabric-react/lib/Modal';

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px'
});

export interface ISPSearchPanelSubSiteListState {
  items: ISPSite[];
  selectionDetails: string;
  results: string;
}
export interface ISPSearchPanelSubSiteListProps {
    sites: ISPSite[];
    updated: boolean;
 
}

export class SPSearchPanelSubSiteList extends React.Component<ISPSearchPanelSubSiteListProps, ISPSearchPanelSubSiteListState> {
  private _selection: Selection;

  private _columns: IColumn[];

  constructor(props: ISPSearchPanelSubSiteListProps) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
    });

    this._columns = [
      { key: 'column1', name: 'Title', fieldName: 'title', minWidth: 40, maxWidth: 120, isResizable: true, onRender: (item: ISPSite) =>{ return <span>{item.Title}</span>;}},
      { key: 'column1', name: 'Relative Url', fieldName: 'url', minWidth: 60, maxWidth: 300, isResizable: true,onRender: (item: ISPSite) =>{ return <span>{item.Url}</span>;} },
      { key: 'column2', name: 'Views Recent', fieldName: 'viewsrecent', minWidth: 60, maxWidth: 100, isResizable: true,onRender: (item: ISPSite) =>{ return <span>{item.ViewsRecent}</span>;} },
      { key: 'column3', name: 'Date User Last Modified', fieldName: 'dateuserlastmodified', minWidth: 120, maxWidth: 150, isResizable: true, onRender: (item: ISPSite) =>{ return <span>{item.LastItemUserModifiedDateFomatted}</span>;}}
    ];
    

    this.state = {
      items: this.props.sites,
      selectionDetails: this._getSelectionDetails(),
      results: "No Sub Webs"
    };
  }

  public componentDidUpdate(){
    if(this.state.results === "No Sub Webs" && this.props.updated && this.props.sites.length > 0){
        let _results: string = "";
        this.setState({
            results: _results
        });
    }
  }


  public render(): JSX.Element {
    const { items, selectionDetails } = this.state;

    return (
      <Fabric>
        <Label>Sub Webs</Label>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            compact={true}
            items={this.props.sites}
            columns={this._columns}
            selectionMode={SelectionMode.none}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={this._onItemInvoked}
            checkButtonAriaLabel="Row checkbox"
          />
        </MarqueeSelection>
        <Label disabled={true}>{this.state.results}</Label>
        
      </Fabric>
    );
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as ISPSite).Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onItemInvoked(item: ISPSite): void {
    alert(`Item invoked: ${item.Title}`);
  }
}