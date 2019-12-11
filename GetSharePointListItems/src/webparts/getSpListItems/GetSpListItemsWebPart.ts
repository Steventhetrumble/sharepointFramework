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

import styles from './GetSpListItemsWebPart.module.scss';
import * as strings from 'GetSpListItemsWebPartStrings';
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
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration, ISPHttpClientOptions, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import * as queryBody1 from './query.json';
import * as queryHeaders1 from './headers.json';
// import {button} from '@microsoft/sp-office-ui-fabric-core';



export interface IGetSpListItemsWebPartProps {
  description: string;
}


export default class GetSpListItemsWebPart extends BaseClientSideWebPart<IGetSpListItemsWebPartProps> {

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


  private _postListData(_url: string = this.context.pageContext.web.absoluteUrl): void {
    const spOpts: ISPHttpClientOptions = {
      headers: queryHeaders1.default,
      body: JSON.stringify(queryBody1.default)
    };
    this.context.spHttpClient.post(`${_url}/_api/search/postquery`, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: Result) => {

          console.log(responseJSON.PrimaryQueryResult.RelevantResults.Table);
        }).catch(err => {
          console.log("error on post", err);
        });
      });
  }

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

  private _getSiteListing(_url: string, rowStart: number): Promise<RelevantResults> {
    // &selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate%27
    return this.context.spHttpClient.get(_url + `/_api/search/query?querytext=%27contentclass:STS_Site Path:https://rcirogers.sharepoint.com/sites/* OR Path:https://rcirogers.sharepoint.com/teams/* %27&trimduplicates=false&RowLimit=500&startrow=${rowStart.toString()}&selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate%27`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: Result) => {
          console.log(responseJSON);
          return responseJSON.PrimaryQueryResult.RelevantResults;
        });
      });
  }

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

  private _renderSortButtons(state: State): void {
    const ButtonContainer: Element = this.domElement.querySelector("#SortButtons");
    ButtonContainer.innerHTML = `<button class="${styles.button}" id="btnSortByLastItem" >Sort by Last Item mod</button><button class="${styles.button}" id="btnSortByRecentViews">Sort by least Recent Views</button>`;
    this.domElement.querySelector(`#btnSortByLastItem`).addEventListener('click', () => {
      state.sortSitesByLastItemUserModified();
      this._renderISPList(state);
    });
    this.domElement.querySelector(`#btnSortByRecentViews`).addEventListener('click', () => {
      state.sortSitesByRecentViews();
      this._renderISPList(state);
    });
  }

  private _createSiteList(state: State, items: Table, totalRows: number): void {
    for (var i = 0; i < Math.floor(totalRows / 500); i++) {
      this._getSiteListing(this.context.pageContext.web.absoluteUrl, 500 + i * 500).then((response: RelevantResults) => {
        response.Table.Rows.forEach((item: Row) => {

          this._getSPDate(item).then((responseDate: ISPDate) => {
            let tempSite: ISPSite;
            if (null != responseDate.value) {
              tempSite = {
                DocID: item.Cells[1].Value,
                Title: item.Cells[2].Value,
                Url: item.Cells[3].Value,
                ViewsLifeTime: Number(item.Cells[6].Value),
                ViewsRecent: Number(item.Cells[7].Value),
                Size: Number(item.Cells[8].Value),
                SiteDescription: item.Cells[9].Value,
                LastItemUserModifiedDate: state.convertSPDate(responseDate.value),
                renderTemplateId: item.Cells[16].Value
              };

            }
            else {
              tempSite = {
                DocID: item.Cells[1].Value,
                Title: item.Cells[2].Value,
                Url: item.Cells[3].Value,
                ViewsLifeTime: Number(item.Cells[6].Value),
                ViewsRecent: Number(item.Cells[7].Value),
                Size: Number(item.Cells[8].Value),
                SiteDescription: item.Cells[9].Value,
                LastItemUserModifiedDate: null,
                renderTemplateId: item.Cells[16].Value
              };

            }

            state.addSite(tempSite);
            if (totalRows === state.getSites().length) {
              this._renderSortButtons(state);
            }
          });
        });
      });
    }

    items.Rows.forEach((item: Row) => {
      this._getSPDate(item).then((responseDate: ISPDate) => {
        let tempSite: ISPSite;
        let html: string = '';
        if (null != responseDate.value) {

          tempSite = {
            DocID: item.Cells[1].Value,
            Title: item.Cells[2].Value,
            Url: item.Cells[3].Value,
            ViewsLifeTime: Number(item.Cells[6].Value),
            ViewsRecent: Number(item.Cells[7].Value),
            Size: Number(item.Cells[8].Value),
            SiteDescription: item.Cells[9].Value,
            LastItemUserModifiedDate: state.convertSPDate(responseDate.value),
            renderTemplateId: item.Cells[16].Value
          };
          html += responseDate.value;

        }
        else {
          tempSite = {
            DocID: item.Cells[1].Value,
            Title: item.Cells[2].Value,
            Url: item.Cells[3].Value,
            ViewsLifeTime: Number(item.Cells[6].Value),
            ViewsRecent: Number(item.Cells[7].Value),
            Size: Number(item.Cells[8].Value),
            SiteDescription: item.Cells[9].Value,
            LastItemUserModifiedDate: null,
            renderTemplateId: item.Cells[16].Value
          };
          html += responseDate.error.message;
        }
        state.addSite(tempSite);
        const DateInfo: Element = this.domElement.querySelector(`#LoadLastItemUserModifiedDate${item.Cells[1].Value}`);
        DateInfo.innerHTML = html;
        if (totalRows === state.getSites().length) {
          this._renderSortButtons(state);
        }
      });
    });
  }

  

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


  public _loadUser(Url: string, DocId: string): void {

    this.context.spHttpClient.get(Url + `/_api/Web/SiteGroups/?$filter=((id eq 5) or (id eq 6))'`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {

        let html: string = '';
        response.json().then((responseJSON: SPGroupList) => {
          if (null != responseJSON.value) {
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
              console.log(responseJSON.value[0]);
              html+=`${responseJSON.value[0].OwnerTitle}:  `;
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


  private _preRender(items: Table): void {
    let html: string = `<div class="${styles.row}"><table class ="Tftable" border = 1 style="border-collapse: collapse;">`;
    let _wanted: wanted = { "Rank": false, "DocId": false, "Title": true, "Url": true, "OriginalPath": false, "Path": false, "ViewsLifeTime": true, "ViewsRecent": true, "Size": true, "SiteDescription": true, "LastItemUserModifiedDate": true, "_ranking_features_": false, "PartitionId": false, "UrlZone": false, "Culture": false, "ResultTypeId": false, "RenderTemplateId": true };
    items.Rows[0].Cells.forEach((_cell: Cell) => {
      if (_wanted[_cell.Key]) {html += `<th>${_cell.Key}</th>`;}
    });
    html += `<th> Load Users </th>`;
    items.Rows.forEach((item: Row) => {
      html += `<tr>`;
      item.Cells.forEach((_cell: Cell) => {
        if (_wanted[_cell.Key]) {
          if (_cell.Key === "LastItemUserModifiedDate") { html += `<td id="LoadLastItemUserModifiedDate${item.Cells[1].Value}"></td>`;}
          else {html += `<td>${_cell.Value}</td>`;}
        }
      });
      html += `<td id="LoadUserFor${item.Cells[1].Value}"><button id="btn${item.Cells[1].Value}" class="${styles.button}">Load Users</button></td></tr>`;
    });
    html += `</table></div>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
    items.Rows.forEach((item: Row) => {
      this.domElement.querySelector(`#btn${item.Cells[1].Value}`).addEventListener('click', () => {
        this._loadUser(item.Cells[3].Value, item.Cells[1].Value);
      });
    });

  }
  private _renderISPList(state: State): void {
    let html: string = `<div class="${styles.row}"><table class ="Tftable" border = 1 style="border-collapse: collapse;">`;
    html += `<th>Title</th><th>Url</th><th>ViewsLifeTime</th><th>ViewsRecent</th><th>Size</th><th>SiteDesciption</th><th>LastItemUserModifiedDate</th><th>Render Template</th><th> Load Users </th>`;


    state.getSearchResults().forEach((site: ISPSite) => {
      html += `<tr><td>${site.Title}</td><td>${site.Url}</td><td>${site.ViewsLifeTime}</td><td>${site.ViewsRecent}</td><td>${site.Size}</td><td>${site.SiteDescription}</td><td>${site.LastItemUserModifiedDate}</td><td>${site.renderTemplateId}</td><td id="LoadUserFor${site.DocID}"><button id="btn${site.DocID}" class="${styles.button}">Load Users</button></td></tr>`;
    });

    html += `</table></div>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
    state.getSearchResults().forEach((site: ISPSite) => {
      this.domElement.querySelector(`#btn${site.DocID}`).addEventListener('click', () => {
        this._loadUser(site.Url, site.DocID);
      });
    });
  }

  public render(): void {

    let state = new State(this._getRootSite());
    this.domElement.innerHTML =
      `<div class="${styles.getSpListItems}">  
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
        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
        <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Site Details</div>  
        <div class="ms-Grid-Row ms-bgColor-themeDark ${styles.row}">
          <div id="SortButtons"><button class="${styles.button}" disabled>Sort by Last Item mod</button><button class="${styles.button}" disabled>Sort by least Recent Views</button></div>
          <div ><button class="${styles.button}" disabled><</button><button class="${styles.button}" disabled>></button></div>
        </div>
        <div id="spListContainer" />  
        </div>  
      </div>  
    </div>`;
    this._renderListAsync(state);
    // var ButtonElements = document.querySelectorAll(".ms-Button");
    // for(var i = 0; i < ButtonElements.length; i++){
    //   new fabric['Button'];
    // }
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
