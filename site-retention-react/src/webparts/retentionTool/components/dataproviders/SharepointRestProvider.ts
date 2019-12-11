import {
  IList,
  ISPSite,
  Usage,
  RelevantResults,
  Result,
  Table,
  Row,
  ListResult,
  ISPDate,
  ListRow,
  SPGroupList,
  ISPGroupEmails,
  ISPUsers,
  SPGroup,
  ISPUser,
  userID
} from "./../common/IObjects";
import { ISharepointRestProvider } from "./ISharepointRestProvider.types";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption, DropdownMenuItemType } from "office-ui-fabric-react";
import {
  ListRows,
  TargetedSites
} from "../RetentionPivot/FindSitesList/FindSitesDetailList/FindSitesDetailList.types";

export class SharepointRestProvider implements ISharepointRestProvider {
  private context: WebPartContext;

  constructor(webContext: WebPartContext) {
    this.context = webContext;
  }

  public getSubWebs(target: ISPSite): Promise<RelevantResults> {
    return this.context.spHttpClient
      .get(
        target.Url +
          `/_api/search/query?querytext=%27(contentclass:STS_Web) Path:${target.Url}/* NOT (WebTemplate:GROUP)%27&trimduplicates=false&selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate,WebTemplate%27`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: Result) => {
          return responseJSON.PrimaryQueryResult.RelevantResults;
        });
      });
  }

  public lookupDates(res: RelevantResults): Promise<ISPSite>[] {
    const randomDate = _randomDate(new Date(2012, 0, 1), new Date(2012, 0, 1));
    let proms: Promise<ISPSite>[] = [];
    res.Table.Rows.forEach((row: Row) => {
      proms.push(
        this.getSPDate(row).then((responseDate: ISPDate) => {
          return {
            Title: row.Cells[2].Value,
            Url: row.Cells[3].Value,
            ViewsLifeTime: Number(row.Cells[6].Value),
            ViewsRecent: Number(row.Cells[7].Value),
            Size: Number(row.Cells[8].Value),
            SiteDescription: row.Cells[9].Value,
            LastItemUserModifiedDateSharepoint:
              responseDate.error == null
                ? responseDate.value
                : "randomDate.date",
            LastItemUserModifiedDate:
              responseDate.error == null ? responseDate.date : randomDate.date,
            LastItemUserModifiedDatevalue:
              responseDate.error == null
                ? responseDate.datevalue
                : randomDate.value,
            LastItemUserModifiedDateFomatted:
              responseDate.error == null
                ? responseDate.dateFormatted
                : randomDate.dateFormatted,
            renderTemplateId: row.Cells[16].Value
          };
        })
      );
    });

    return proms;
  }

  public getUserGroups(target: ISPSite): Promise<SPGroupList> {
    return this.context.spHttpClient
      .get(target.Url + `/_api/Web/SiteGroups`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: SPGroupList) => {
          return responseJSON;
        });
      });
  }

  public deleteSiteFromList = (siteSelected: ListRow): Promise<boolean> => {
    console.log(siteSelected["@odata.editLink"]);
    return this.context.spHttpClient
      .post(
        `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/${siteSelected["@odata.editLink"]}`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
          }
        }
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return true;
        }
      })
      .catch(err => {
        return false;
      });
  };

  public getRootSiteData(): Promise<ISPSite> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl + "/_api/site/usage",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json().then((responseJSON: Usage) => {
            const site: ISPSite = {
              Title: this.context.pageContext.web.title,
              Url: this.context.pageContext.web.absoluteUrl,
              ViewsLifeTime: responseJSON.ViewsLifeTime,
              ViewsRecent: responseJSON.ViewsRecent,
              Size: responseJSON.Storage,
              SiteDescription: this.context.pageContext.web.description,
              LastItemUserModifiedDateSharepoint: null,
              LastItemUserModifiedDate: null,
              LastItemUserModifiedDateFomatted: null,
              LastItemUserModifiedDatevalue: null,
              renderTemplateId: this.context.pageContext.web.templateName
            };

            return site;
          });
        } else {
          console.log(response.status);
          return null;
        }
      });
  }

  public getTotalSites(): Promise<number> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/search/query?querytext=%27(contentclass:STS_Site) Path:https://rcirogers.sharepoint.com/sites/* OR Path:https://rcirogers.sharepoint.com/teams/* NOT (WebTemplate:GROUP) %27&trimduplicates=false&RowLimit=1&startrow=0&selectproperties=%27Title%27`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: Result) => {
          return responseJSON.PrimaryQueryResult.RelevantResults.TotalRows;
        });
      });
  }

  public getEmails(): Promise<string[]> {
    return this.context.spHttpClient
      .get(
        `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetedEmailTemplates')/Items?$top=1000`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json().then((responseJSON: ListResult) => {
            let emails: string[];
            responseJSON.value.forEach((element: ListRow) => {
              emails.push(element.emails);
            });
            return emails;
          });
        }
      });
  }
  public saveEmail(contents: string): void {
    var itemType = getItemTypeForListName("TargetedEmailTemplates");

    const body: string = JSON.stringify({
      __metadata: {
        type: itemType
      },
      bodyTemplate: contents
    });
    this.context.spHttpClient
      .post(
        `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/Lists(guid'01cb56d5-817b-4358-a1c8-b328342534e6')/Items(1)`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
            Accept: "application/json;odata=verbose",
            "Content-type": "application/json;odata=verbose",
            "odata-version": ""
          },
          body: body
        }
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log("Save response:", response.status);
        }
      });
  }
  public saveTarget(targetedSite: ListRow): void {
    var itemType = "SP.Data.TargetSitesForDeletionSteven1234ListItem";
    const body: string = JSON.stringify({
      __metadata: {
        type: itemType
      },
      FirstEmailSent: targetedSite.FirstEmailSent,
      FirstEmailReply: targetedSite.FirstEmailReply,
      SecondEmailSent: targetedSite.SecondEmailSent,
      SecondEmailReply: targetedSite.SecondEmailReply,
      Deleted: targetedSite.Deleted,
      DeletionDate: targetedSite.DeletionDate
    });
    this.context.spHttpClient
      .post(
        `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/` +
          targetedSite["@odata.editLink"],
        SPHttpClient.configurations.v1,
        {
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
            Accept: "application/json;odata=verbose",
            "Content-type": "application/json;odata=verbose",
            "odata-version": ""
          },
          body: body
        }
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log("Save response:", response.status);
        }
      });
  }

  public getTargetedSitesListing(rowStart?: number): Promise<ListResult> {
    return this.context.spHttpClient
      .get(
        `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetSitesForDeletionSteven1234')/Items?$top=1000`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: ListResult) => {
          return responseJSON;
        });
      });
  }

  public getSitesListing(rowStart?: number): Promise<Row[]> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/search/query?querytext=%27(contentclass:STS_Site) Path:https://rcirogers.sharepoint.com/sites/* OR Path:https://rcirogers.sharepoint.com/teams/* NOT (WebTemplate:GROUP) %27&trimduplicates=false&RowLimit=500&startrow=${rowStart.toString()}&selectproperties=%27Title,Url,Path,ViewsLifeTime,ViewsRecent,Size,SiteDescription,LastItemUserModifiedDate,WebTemplate%27`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: Result) => {
          return responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows;
        });
      });
  }

  public postListData(site: ISPSite, users: ISPUser[]): void {
    console.log(site.LastItemUserModifiedDateSharepoint);
    var itemType = getItemTypeForListName("TargetSitesForDeletionSteven1234");

    let targetedUsers: Promise<number>[] = [];

    users.forEach((user: ISPUser) => {
      if (user.Title != null) {
        targetedUsers.push(this.getByEmail(user.Email));
      }
    });

    Promise.all(targetedUsers).then(responses => {
      let finalTargets: number[] = [];

      responses.forEach(response => {
        if (response > 0) {
          finalTargets.push(response);
        }
      });
      const body: string = JSON.stringify({
        __metadata: {
          type: itemType
        },
        Title: site.Title,
        URL: site.Url,
        UserLastModified: site.LastItemUserModifiedDateSharepoint,
        RecentViews: site.ViewsRecent,
        peopleId: { results: finalTargets }
      });
      // make the actuall post request and handle the resulsts
      this.context.spHttpClient
        .post(
          `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetSitesForDeletionSteven1234')/Items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-type": "application/json;odata=verbose",
              "odata-version": ""
            },
            body: body
          }
        )
        .then((response: SPHttpClientResponse) => {
          response
            .json()
            .then((responseJSON: Result) => {
              console.log(responseJSON);
            })
            .catch(err => {
              console.log("error on post", err);
            });
        });
    });
  }

  public makeOptions(
    groups: SPGroupList,
    url: string
  ): Promise<[IDropdownOption[], { [id: number]: ISPUsers }]> {
    return new Promise<[IDropdownOption[], { [id: number]: ISPUsers }]>(
      resolve => {
        let _options: IDropdownOption[] = [];
        let _all_users: { [id: number]: ISPUsers } = {};
        _options.push({
          key: "SuggestedOptionsHeader",
          text: "SuggestedGroup",
          itemType: DropdownMenuItemType.Header
        });
        let bestOption: IDropdownOption = {
          key: "NoOptions",
          text: "No Suggestions",
          disabled: true
        };

        groups.value.forEach((group: SPGroup) => {
          if (group.Id == 5) {
            bestOption = {
              key: group.Id,
              text: `ID:${group.Id} ${group.OwnerTitle}`
            };
          } else if (group.Id == 6 && bestOption.key != 5) {
            bestOption = {
              key: group.Id,
              text: `ID:${group.Id} ${group.OwnerTitle}`
            };
          }
        });

        _options.push(bestOption);
        _options.push({
          key: "groups",
          text: "Group Options",
          itemType: DropdownMenuItemType.Header
        });

        groups.value.forEach((group: SPGroup) => {
          if (group.Id != bestOption.key) {
            if (
              group.OwnerTitle === "System Account" ||
              group.OwnerTitle === "serv_epmwss"
            ) {
              _options.push({
                key: group.Id,
                text: `ID:${group.Id} ${group.OwnerTitle}`,
                disabled: true
              });
            } else {
              this.context.spHttpClient
                .get(
                  url + `/_api/Web/SiteGroups/GetById(${group.Id})/Users`,
                  SPHttpClient.configurations.v1
                )
                .then((response: SPHttpClientResponse) => {
                  response.json().then((responseJSON: ISPUsers) => {
                    if (
                      responseJSON.error == null &&
                      responseJSON.OwnerTitle != null
                    ) {
                      _all_users[group.Id] = responseJSON;
                      _all_users[group.Id].OwnerTitle = group.OwnerTitle;
                      _options.push({
                        key: group.Id,
                        text: `ID:${group.Id} ${group.OwnerTitle}`
                      });
                    } else {
                      console.log(responseJSON.error);
                      _options.push({
                        key: group.Id,
                        text: `ID:${group.Id} ${group.OwnerTitle}`,
                        disabled: true
                      });
                    }
                  });
                })
                .catch(err => {
                  console.log(err);
                });
            }
          } else {
            this.context.spHttpClient
              .get(
                url + `/_api/Web/SiteGroups/GetById(${group.Id})/Users`,
                SPHttpClient.configurations.v1
              )
              .then((response: SPHttpClientResponse) => {
                response.json().then((responseJSON: ISPUsers) => {
                  if (responseJSON.error == null) {
                    let _users: ISPUsers = responseJSON;
                    _all_users[group.Id] = _users;
                    _all_users[group.Id].OwnerTitle = group.OwnerTitle;
                  }
                });
              });
          }
        });
        return resolve([_options, _all_users]);
      }
    );
  }

  private _getTargetedSites(): Promise<TargetedSites> {
    return this.context.spHttpClient
      .get(
        `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetSitesForDeletionSteven1234')/Items?$top=1000`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responesJSON: ListRows) => {
          let _targetedSites: TargetedSites = {};
          responesJSON.value.forEach((row: ListRow) => {
            _targetedSites[row.Title] = true;
          });

          return _targetedSites;
        });
      });
  }
  public getByEmail(email: string): Promise<number> {
    email = email.toLocaleLowerCase();
    return this.context.spHttpClient
      .get(
        `https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/SiteUsers/getByEmail('${email}')/Id`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json().then((resID: userID) => {
            return resID.value;
          });
        } else {
          return -1;
        }
      })
      .catch(err => {
        console.log(err);
        return -1;
      });
  }

  public getSPDate(row: Row): Promise<ISPDate> {
    return this.context.spHttpClient
      .get(
        row.Cells[3].Value + `/_api/Web/LastItemUserModifiedDate`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.status != 200) {
          console.log(response.status);
          console.log(response.statusText);
        }
        return response
          .json()
          .then((responseJSON: ISPDate) => {
            if (responseJSON.error == null) {
              let _date: Date = this._convertSPDate(responseJSON.value);
              return {
                value: responseJSON.value,
                dateFormatted: _date.toLocaleDateString(),
                datevalue: _date.valueOf(),
                error: null
              };
            } else {
              return {
                value: null,
                DateFormatted: null,
                datevalue: null,
                error: responseJSON.error
              };
            }
          })
          .catch(err => {
            console.log(`convert to json error in SPDATE: ${err}`);
            return err;
          });
      })
      .catch(err => {
        console.log(`get request error at SPDATE: ${err}`);
        return err;
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
    let dDate: Date = new Date(
      Number(xYear),
      Number(xMonth) - 1,
      Number(xDay),
      Number(xHour),
      Number(xMin),
      Number(xSec)
    );
    return dDate;
  }

  public getAllLists(): Promise<IList[]> {
    let _items: IList[];
    //Initiate mockup values to the IList[] object
    _items = [
      {
        Title: "List Name 1",
        Id: "1"
      },
      {
        Title: "List Name 2",
        Id: "2"
      },
      {
        Title: "List Name 3",
        Id: "3"
      },
      {
        Title: "List Name 4",
        Id: "4"
      },
      {
        Title: "List Name 5",
        Id: "5"
      }
    ];

    //Returns the mockup data
    return new Promise<IList[]>(resolve => {
      setTimeout(() => {
        resolve(_items);
      }, 2000);
    });
  }
}

function _randomDate(
  start: Date,
  end: Date
): { value: number; dateFormatted: string; date: Date } {
  const _date: Date = new Date(
    start.getTime() + Math.random() * (end.getTime() - start.getTime())
  );
  return {
    date: _date,
    value: _date.valueOf(),
    dateFormatted: _date.toLocaleDateString()
  };
}

// Get List Item Type metadata
function getItemTypeForListName(name) {
  return (
    "SP.Data." +
    name.charAt(0).toUpperCase() +
    name
      .split(" ")
      .join("")
      .slice(1) +
    "ListItem"
  );
}
