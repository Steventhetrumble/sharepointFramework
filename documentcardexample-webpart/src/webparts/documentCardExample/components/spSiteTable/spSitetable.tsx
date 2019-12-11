import * as React from 'react';
import { getRTL } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { List } from 'office-ui-fabric-react/lib/List';
import { ITheme, mergeStyleSets, getTheme, getFocusStyle } from 'office-ui-fabric-react/lib/Styling';
import { createListItems, IExampleItem } from 'office-ui-fabric-react/lib/utilities/exampleData';
import {
    ISPLists,
    ISPList,
    Row,
    Table,
    RelevantResults,
    PrimaryQueryResult,
    Cell,
    Result,
    wanted,
    ISPDate,
    ISPUsers,
    ISPUser,
    State,
    ISPSite,
    Usage,
    SPGroupList,
    SPGroup
  } from '../GetListItemsWebPart';


export class spSitetable extends React.Component {
  

  constructor(props: State) {
    super(props);
    this.state = {
      sites: "",
      items: ""
    };
  }




  public render(): JSX.Element {

    return (
        <table className= "ms-Table">
            <thead>
                <tr>
                    <th>Location</th>
                    <th>Modified</th>
                    <th>Type</th>
                    <th>File Name</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Location</td>
                    <td>Modified</td>
                    <td>Type</td>
                    <td>File Name</td>
                </tr>
                <tr>
                    <td>Location</td>
                    <td>Modified</td>
                    <td>Type</td>
                    <td>File Name</td>
                </tr>
                <tr>
                    <td>Location</td>
                    <td>Modified</td>
                    <td>Type</td>
                    <td>File Name</td>
                </tr>
                <tr>
                    <td>Location</td>
                    <td>Modified</td>
                    <td>Type</td>
                    <td>File Name</td>
                </tr>
                <tr>
                    <td>Location</td>
                    <td>Modified</td>
                    <td>Type</td>
                    <td>File Name</td>
                </tr>
                <tr>
                    <td>Location</td>
                    <td>Modified</td>
                    <td>Type</td>
                    <td>File Name</td>
                </tr>
                <tr>
                    <td>Location</td>
                    <td>Modified</td>
                    <td>Type</td>
                    <td>File Name</td>
                </tr>
            </tbody>
        </table>
     
    );
  }
}
