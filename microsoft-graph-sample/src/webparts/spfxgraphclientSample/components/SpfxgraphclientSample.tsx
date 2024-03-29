import * as React from 'react';
import styles from './SpfxgraphclientSample.module.scss';
import { ISpfxgraphclientSampleProps } from './ISpfxgraphclientSampleProps';
import * as strings from 'SpfxgraphclientSampleWebPartStrings';
import { ISpfxgraphclientSampleState } from './ISpfxgraphclientSampleState';
import { ClientMode } from './ClientMode';
import { IUserItem } from './IUserItem';
import { escape } from '@microsoft/sp-lodash-subset';
import { any } from 'prop-types';
import {
  autobind,
  PrimaryButton,
  TextField,
  Label,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode
} from 'office-ui-fabric-react';
import { AadHttpClient, MSGraphClient} from "@microsoft/sp-http";

let _usersListColumns = [
  {
    key: 'displayName',
    name: 'Display name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'mail',
    name: 'Mail',
    fieldName: 'mail',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'userPrincipalName',
    name: 'User Principal Name',
    fieldName: 'userPrincipalName',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
];

export default class SpfxgraphclientSample extends React.Component<ISpfxgraphclientSampleProps, ISpfxgraphclientSampleState> {
  constructor(props: ISpfxgraphclientSampleProps){
    super(props);

    this.state = {
      users: [],
      searchFor: ""
    }


    

  }


  


  public render(): React.ReactElement<ISpfxgraphclientSampleProps> {
    const { clientMode } = this.props;
    return (
      <div className={ styles.spfxgraphclientSample }>
      <div className={ styles.container }>
        <div className={ styles.row }>
          <div className={ styles.column }>
            <span className={ styles.title }>Search for a user!</span>
            <p className={ styles.form }>
              <TextField 
                  label={ strings.SearchFor } 
                  required={ true } 
                  value={ this.state.searchFor }
                  onChanged={ this._onSearchForChanged }
                  onGetErrorMessage={ this._getSearchForErrorMessage }
                />
            </p>
            {
              (clientMode === ClientMode.aad || clientMode === ClientMode.graph) ?
                <p className={styles.form}>
                  <PrimaryButton
                    text='Search'
                    title='Search'
                    onClick={this._search}
                  />
                </p>
                : <p>Configure client mode by editing web part properties.</p>
            }
            {
              (this.state.users != null && this.state.users.length > 0) ?
                <p className={ styles.form }>
                <DetailsList
                    items={ this.state.users }
                    columns={ _usersListColumns }
                    setKey='set'
                    checkboxVisibility={ CheckboxVisibility.hidden }
                    selectionMode={ SelectionMode.none }
                    layoutMode={ DetailsListLayoutMode.fixedColumns }
                    compact={ true }
                />
              </p>
              : null
            }
          </div>
        </div>
      </div>
    </div>
  );
  }

  @autobind
  private _onSearchForChanged(newValue: string): void {

    //update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue,
    });
  }

  private _getSearchForErrorMessage(value: string): string {
    // The search for text cannot contain spaces
    return (value == null || value.length == 0 || value.indexOf(" ") <  0)
    ? ''
    : `${strings.SearchForValidationErrorMessage}`;
  }

  @autobind
  private _search(): void {

    // Based on the clienMode value search users
    switch (this.props.clientMode)
    {
      case ClientMode.aad:
        this._searchWithAad();
        break;
      case ClientMode.graph:
        this._searchWithGraph();
        break;
    }
  }

  private _searchWithGraph(): void {
    // Log the current operation
  console.log("Using _searchWithGraph() method");

  this.props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
      client
        .api("users/delta")
        .version("v1.0")
        .get((err, res) => {  

          if (err) {
            console.error(err);
            return;
          }

          // Prepare the output array
          var users: Array<IUserItem> = new Array<IUserItem>();
          console.log(res);
          // Map the JSON response to the output array
          res.value.map((item: any) => {
            users.push( { 
              displayName: item.displayName,
              mail: item.mail,
              userPrincipalName: item.userPrincipalName,
            });
          });

          // Update the component state accordingly to the result
          this.setState(
            {
              users: users,
            }
          );
        });
    });
  }

  private _searchWithAad(): void{
    // Log the current operation
  console.log("Using _searchWithAad() method");

  // Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
  this.props.context.aadHttpClientFactory
    .getClient('https://graph.microsoft.com')
    .then((client: AadHttpClient) => {
      // Search for the users with givenName, surname, or displayName equal to the searchFor value
      return client
        .get(
          `https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName&$filter=(givenName%20eq%20'${escape(this.state.searchFor)}')%20or%20(surname%20eq%20'${escape(this.state.searchFor)}')%20or%20(displayName%20eq%20'${escape(this.state.searchFor)}')`,
          AadHttpClient.configurations.v1
        );
    })
    .then(response => {
      let res = response.json();
      console.log(res);
      return res;
    })
    .then(json => {

      // Prepare the output array
      var users: Array<IUserItem> = new Array<IUserItem>();

      // Log the result in the console for testing purposes
      console.log(json);

      // Map the JSON response to the output array
      json.value.map((item: any) => {
        users.push( {
          displayName: item.displayName,
          mail: item.mail,
          userPrincipalName: item.userPrincipalName,
        });
      });

      // Update the component state accordingly to the result
      this.setState(
        {
          users: users,
        }
      );
    })
    .catch(error => {
      console.error(error);
    });

  }
}


