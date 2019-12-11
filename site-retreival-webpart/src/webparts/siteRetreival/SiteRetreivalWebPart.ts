import {
  Version,
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SiteRetreivalWebPart.module.scss';
import * as strings from 'SiteRetreivalWebPartStrings';
import MockHttpClient from './MockHttpClient';
import MockHttpClient2 from './MockHttpClient2';
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
} from './GetListItemsWebPart';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration, ISPHttpClientOptions, ISPHttpClientConfiguration, ISPHttpClientBatchCreationOptions, SPHttpClientBatch } from '@microsoft/sp-http';
import * as queryBody1 from './query.json';
import * as queryHeaders1 from './headers.json';
import { SPUser } from '@microsoft/sp-page-context';

export interface ISiteRetreivalWebPartProps {
  description: string;
}

export default class SiteRetreivalWebPart extends BaseClientSideWebPart<ISiteRetreivalWebPartProps> {
  // mock list data for testing
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {
      const listData: ISPLists = {
        value:
          [
            { 'EmployeeId': 'E123', 'EmployeeName': 'John', 'Experience': 'SharePoint', 'Location': 'India' },
            { 'EmployeeId': 'E567', 'EmployeeName': 'Martin', 'Experience': '.Net', 'Location': 'Qatar' },
            { 'EmployeeId': 'E367', 'EmployeeName': 'Luke', 'Experience': 'Java', 'Location': 'Uk' },
          ]
      };
      return listData;
    }) as Promise<ISPLists>;
  }


  // get mock data to fill out the local version of the app. 
  private _getMockListData2(): Promise<ISPSite> {
    return MockHttpClient2.get(this.context.pageContext.web.absoluteUrl).then(() => {
      const SiteData: ISPSite = {
        DocID: "0",
        Title: "local",
        Url: "local",
        ViewsLifeTime: 3,
        ViewsRecent: 2,
        Size: 24,
        SiteDescription: "this is a local site and can not be determined",
        LastItemUserModifiedDate: new Date(),
        renderTemplateId: "templateid"
      };
      return SiteData;
    }) as Promise<ISPSite>;
  }

  // render list for mock data
  private _renderList(items: ISPList[]): void {
    let html: string = '<table class ="Tftable" border = 1 width = 100% style="border-collapse: collapse;">';
    html += `<th>EmployeeId</th><th>EmplyeeName</th><th>Experience</th><th>Location</th>`;
    items.forEach((item: ISPList) => {
      html += `
        <tr>
        <td>${item.EmployeeId}</td>
        <td>${item.EmployeeName}</td>
        <td>${item.Experience}</td>
        <td>${item.Location}</td>
        </tr>
      `;

    });
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  // and example of a post requrest... perhaps will need this at some point
  private _postListData(_url: string = this.context.pageContext.web.absoluteUrl): void {
    // the following options are each defined in their own json files
    const spOpts: ISPHttpClientOptions = {
      headers: queryHeaders1.default,
      body: JSON.stringify(queryBody1.default)
    };
    // make the actuall post request and handle the resulsts
    this.context.spHttpClient.post(`${_url}/_api/search/postquery`, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: Result) => {

          console.log(responseJSON.PrimaryQueryResult.RelevantResults.Table);
        }).catch(err => {
          console.log("error on post", err);
        });
      });
  }

  // this is to get the results of the tenant.  this should give all of the information regarding the site
  // this is used for rendering the title block of the site
  private _getRootSiteData(_url: string = this.context.pageContext.web.absoluteUrl): Promise<ISPSite> {
    return this.context.spHttpClient.get(_url + '/_api/site/usage', SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: Usage) => {
          console.log(responseJSON);
          // TODO views lifetime and views recent 
          const site: ISPSite = {
            DocID: "",
            Title: this.context.pageContext.web.title,
            Url: this.context.pageContext.web.absoluteUrl,
            ViewsLifeTime: 0,
            ViewsRecent: 0,
            Size: responseJSON.Storage,
            SiteDescription: this.context.pageContext.web.description,
            LastItemUserModifiedDate: null,
            renderTemplateId: ""
          };
          return site;
        });
      });
  }

  // the request that should provide all sites belonging to the given tenant.  This is based off of the users context
  // so some sites will be hidden depending on the users permissions.
  private _getSiteListing(_url: string, rowStart: number): Promise<RelevantResults> {
    // &selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate%27
    return this.context.spHttpClient.get(_url + `/_api/search/query?querytext=%27(contentclass:STS_Site  OR contentclass:STS_Web) Path:https://rcirogers.sharepoint.com/sites/* OR Path:https://rcirogers.sharepoint.com/teams/* NOT (WebTemplate:GROUP)%27&trimduplicates=false&RowLimit=500&startrow=${rowStart.toString()}&selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate%27`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: Result) => {
          console.log(responseJSON);
          return responseJSON.PrimaryQueryResult.RelevantResults;
        });
      });
  }

  // this is used to handle the rendering of the title block in the case that the app is being hosted in a local environment
  private _getRootSite(): Promise<ISPSite> {
    if (Environment.type === EnvironmentType.Local) {
      return this._getMockListData2().then((response) => {
        console.log(response);
        return response;
      });
    }
    else {
      return this._getRootSiteData().then((response) => {
        return response;
      });
    }
  }

  // grab the relevant title block elements and inject the resolved values
  private _renderTitleBlock(state: State, TotalResults: number): void {
    const TitleContainer: Element = this.domElement.querySelector('#Title');
    TitleContainer.innerHTML = `Title: ${state.getRootSite().Title}`;
    const DescriptionContainer: Element = this.domElement.querySelector('#Description');
    DescriptionContainer.innerHTML = `Description: ${state.getRootSite().SiteDescription}`;
    const UrlContainer: Element = this.domElement.querySelector('#Url');
    UrlContainer.innerHTML = `Url: ${state.getRootSite().Url}`;
    const SizeContainer: Element = this.domElement.querySelector('#Size');
    SizeContainer.innerHTML = `Size: ${state.getRootSite().Size}`;
    const ResultsContainer: Element = this.domElement.querySelector('#Results');
    ResultsContainer.innerHTML = `Results: ${TotalResults}`;

  }

  // the date must be requested seperately, as it is not provided in the STS_Site querry
  // this means that the information that may be used to sort the items may not be present yet
  // as such, this function will be called when all of the requests have completed
  private _renderSortButtons(state: State): void {
    const ButtonContainer: Element = this.domElement.querySelector("#SortButtons");
    ButtonContainer.innerHTML = `<button class="button" id="btnSortByLastItem" >Sort by Last Item mod</button><button class="button" id="btnSortByRecentViews">Sort by least Recent Views</button>`;
    this.domElement.querySelector(`#btnSortByLastItem`).addEventListener('click', () => {
      state.sortSitesByLastItemUserModified();
      this._renderISPList(state);
    });
    this.domElement.querySelector(`#btnSortByRecentViews`).addEventListener('click', () => {
      state.sortSitesByRecentViews();
      this._renderISPList(state);
    });
  }

  // store all of the site information in an array of objects.
  // these objects must be created, and the date must be requested
  private _createSiteList(state: State, items: Table, totalRows: number): void {
    // a rest call will only respond with a maximum of 500 items but will provide the number of total rows.
    // the rows can then be requested 500 at a time untill all of the sites have been retreived
    for (var i = 0; i < Math.floor(totalRows / 500); i++) {
      this._getSiteListing(this.context.pageContext.web.absoluteUrl, 500 + i * 500).then((response: RelevantResults) => {
        response.Table.Rows.forEach((item: Row) => {

          this._getSPDate(item).then((responseDate: ISPDate) => {
            let tempSite: ISPSite;
            tempSite = {
              DocID: item.Cells[1].Value,
              Title: item.Cells[2].Value,
              Url: item.Cells[3].Value,
              ViewsLifeTime: Number(item.Cells[6].Value),
              ViewsRecent: Number(item.Cells[7].Value),
              Size: Number(item.Cells[8].Value),
              SiteDescription: item.Cells[9].Value,
              LastItemUserModifiedDate: (null!= responseDate.value) ? state.convertSPDate(responseDate.value): null,
              renderTemplateId: item.Cells[16].Value
            };
            state.addSite(tempSite);
            if (totalRows === state.getSites().length) {
              // render the sort button when all of the objects are created
              this._renderSortButtons(state);
            }
          });
        });
      });
    }
    // for the items already retreive. request the date and create the objects
    items.Rows.forEach((item: Row) => {
      this._getSPDate(item).then((responseDate: ISPDate) => {
        let tempSite: ISPSite;
        let html: string = '';
        tempSite = {
          DocID: item.Cells[1].Value,
          Title: item.Cells[2].Value,
          Url: item.Cells[3].Value,
          ViewsLifeTime: Number(item.Cells[6].Value),
          ViewsRecent: Number(item.Cells[7].Value),
          Size: Number(item.Cells[8].Value),
          SiteDescription: item.Cells[9].Value,
          LastItemUserModifiedDate: (null!= responseDate.value) ? state.convertSPDate(responseDate.value): null,
          renderTemplateId: item.Cells[16].Value
        };
        html += (null != responseDate.value) ? responseDate.value : responseDate.error.message;
        state.addSite(tempSite);
        const DateInfo: Element = this.domElement.querySelector(`#LoadLastItemUserModifiedDate${item.Cells[1].Value}`);
        DateInfo.innerHTML = html;
        if (totalRows === state.getSites().length) {
          // render the sort buttons when all the the objects are created
          this._renderSortButtons(state);
        }
      });
    });
  }


  // request the overall site information and then render the page as it is delivered
  // then create objects out of the information and provide other options for navigating/ searching the information
  private _renderListAsync(state: State): void {
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else {
      this._getSiteListing(this.context.pageContext.web.absoluteUrl, 0)
        .then((response) => {
          debugger;
          this._renderTitleBlock(state, response.TotalRows);
          this._preRender(response.Table);
          this._createSiteList(state, response.Table, response.TotalRows);
        }).catch(err => {
          // Do something for an error here
          console.log("Error Reading data " + err);
        });
    }
  }


  // buttons are provided to load users for a site
  // this will request the information and modify the relevant table cell
  public _loadUser(Url: string, DocId: string): void {
    // make a request to see if either of these two groups exist
    // these groups seem to be ownership in a vast majority of sites
    this.context.spHttpClient.get(Url + `/_api/Web/SiteGroups/?$filter=((id eq 5) or (id eq 6))'`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        let html: string = '';
        response.json().then((responseJSON: SPGroupList) => {

          // if the response returns an actual value
          if (null != responseJSON.value) {
            // if the response is an empty array, one of the two groups does not exist
            // so write a message in the appropriate cell, and write the sitegroups object to the console
            // so that an appropriate group can be investigated further
            if (responseJSON.value.length == 0) {
              html += `this group structure needs to be investigated.  Check Console for sitegroups listing.`;
              this.context.spHttpClient.get(Url + `/_api/Web/SiteGroups`,
                SPHttpClient.configurations.v1)
                .then((response2: SPHttpClientResponse) => {
                  response2.json().then((response2JSON: JSON) => {
                    console.log(response2JSON);
                  });
                });
              const userList: Element = this.domElement.querySelector(`#LoadUserFor${DocId}`);
              userList.innerHTML = html;
            }
            else {
              // if there are more then one of the appropriate group, or just one, choose the first and look up the users for that group
              // then take those users and the name of the group and write it to the appropriate cell
              console.log(responseJSON.value[0]);
              html += `${responseJSON.value[0].OwnerTitle}:  `;
              this.context.spHttpClient.get(Url + `/_api/Web/SiteGroups/GetById(${String(responseJSON.value[0].Id)})/Users`,
                SPHttpClient.configurations.v1)
                .then((response2: SPHttpClientResponse) => {
                  response2.json().then((response2JSON: ISPUsers) => {
                    if (response2JSON.value.length == 0) {
                      html += `there are no users in this group`;
                    }
                    response2JSON.value.forEach((item: ISPUser) => {
                      html += `${item.Title};`;
                    });
                    const userList: Element = this.domElement.querySelector(`#LoadUserFor${DocId}`);
                    userList.innerHTML = html;
                  }).catch(err => {

                    console.log(`Load Users: Site group by id: convert to json: ${err} `);
                  });
                });
            }
          }
          else {
            // in the case of an error (403 forbidden for exampel) write the error to the corresponding cell
            html += responseJSON.error.message;
          }
        }).catch(err => {
          console.log("loadUser convert to json error");
          console.log(err);
        });
      }).catch(err => {
        console.log("load user get request error");
        console.log(err);
      });
  }


  // a function to request the last item user modified date from each site
  private _getSPDate(row: Row): Promise<ISPDate> {
    return this.context.spHttpClient.get(row.Cells[3].Value + `/_api/Web/LastItemUserModifiedDate`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status != 200) {
          console.log(response.status);
          console.log(response.statusText);
        }
        return response.json().then((responseJSON: ISPDate) => {
          return responseJSON;
        }).catch(err => {
          console.log(`convert to json error in SPDATE: ${err}`);
          return err;
        });
      }).catch(err => {
        console.log(`get request error at SPDATE: ${err}`);
        return err;
      });
  }

  // when the initial json comes back, render a temporary form summarizing the findings
  // when the json is handled, and the dates are requested(LastitemmodifiedDate), then provide the option to sort and re-render
  private _preRender(items: Table): void {
    // create table
    let html: string = `<div class="${styles.row}"><table class ="Tftable" border = 1 style="border-collapse: collapse;">`;
    // ensure table headings are only the ones desired
    let _wanted: wanted = { "Rank": false, "DocId": false, "Title": true, "Url": true, "OriginalPath": false, "Path": false, "ViewsLifeTime": true, "ViewsRecent": true, "Size": true, "SiteDescription": true, "LastItemUserModifiedDate": true, "_ranking_features_": false, "PartitionId": false, "UrlZone": false, "Culture": false, "ResultTypeId": false, "RenderTemplateId": true };
    items.Rows[0].Cells.forEach((_cell: Cell) => {
      if (_wanted[_cell.Key]) {
        html += `<th>${_cell.Key}</th>`;
      }
    });
    // add load user button
    html += `<th> Load Users </th>`;
    items.Rows.forEach((item: Row) => {
      html += `<tr>`;
      item.Cells.forEach((_cell: Cell) => {
        if (_wanted[_cell.Key]) {
          if (_cell.Key === "LastItemUserModifiedDate") {
            html += `<td id="LoadLastItemUserModifiedDate${item.Cells[1].Value}"></td>`;
          }
          else {
            html += `<td>${_cell.Value}</td>`;
          }
        }
      });
      html += `<td id="LoadUserFor${item.Cells[1].Value}"><button id="btn${item.Cells[1].Value}" class="button">Load Users</button></td></tr>`;
    });
    html += `</table></div>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
    // add an event listener to each button
    items.Rows.forEach((item: Row) => {
      this.domElement.querySelector(`#btn${item.Cells[1].Value}`).addEventListener('click', () => {
        this._loadUser(item.Cells[3].Value, item.Cells[1].Value);
      });
    });

  }

  // the function to render the completed list of sites
  // will be called after the list is sorted 
  private _renderISPList(state: State): void {
    // create table head and labels
    let html: string = `<div class="${styles.row}"><table class ="Tftable" border = 1 style="border-collapse: collapse;">`;
    html += `<th>Title</th><th>Url</th><th>ViewsLifeTime</th><th>ViewsRecent</th><th>Size</th><th>SiteDesciption</th><th>LastItemUserModifiedDate</th><th>Render Template</th><th> Load Users </th>`;

    // add rows
    state.getSearchResults().forEach((site: ISPSite) => {
      html += `<tr><td>${site.Title}</td><td>${site.Url}</td><td>${site.ViewsLifeTime}</td><td>${site.ViewsRecent}</td><td>${site.Size}</td><td>${site.SiteDescription}</td><td>${site.LastItemUserModifiedDate}</td><td>${site.renderTemplateId}</td><td id="LoadUserFor${site.DocID}"><button id="btn${site.DocID}" class="button">Load Users</button></td></tr>`;
    });

    html += `</table></div>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
    // add event listener to each button
    state.getSearchResults().forEach((site: ISPSite) => {
      this.domElement.querySelector(`#btn${site.DocID}`).addEventListener('click', () => {
        this._loadUser(site.Url, site.DocID);
      });
    });
  }

  public render(): void {

    let state = new State(this._getRootSite());
    this.domElement.innerHTML =
      `<div class="${styles.siteRetreival}">  
      <div class="${styles.container}">  
        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
          <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Welcome to SharePoint Framework Development</span>  
          
          <p class="ms-font-l ms-fontColor-white" style="text-align: center">Demo : Retrieve Site Data from SharePoint Rest Api</p>  
          <p class="ms-font-l ms-fontColor-white" style="text-align: center" id="Title">Title of Root Site : </p>  
          <p class="ms-font-l ms-fontColor-white" style="text-align: center" id="Description">Description : </p>  
          <p class="ms-font-l ms-fontColor-white" style="text-align: center" id="Url">Url : </p>  
          <p class="ms-font-l ms-fontColor-white" style="text-align: center" id="Size">Size :</p>  
          <p class="ms-font-l ms-fontColor-white" style="text-align: center" id="Results">Results: </p>  
          
          </div>  
        </div>
        <div class="ms-Grid-Row ms-bgColor-themeDark ${styles.row}" id="SortButtons"><button class="button" disabled>Sort by Last Item mod</button><button class="button" disabled>Sort by least Recent Views</button></div>
        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
        <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Site Details</div>  
        <br>
        <div id="spListContainer" />  
        </div>  
      </div>  
    </div>`;
    this._renderListAsync(state);

    // this._postListData();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
