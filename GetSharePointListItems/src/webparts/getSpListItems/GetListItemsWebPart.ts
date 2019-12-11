export interface ISPLists {
    value: ISPList[];
}
export interface ISPList {
    EmployeeId: string;
    EmployeeName: string;
    Experience: string;
    Location: string;
}

export interface Result {
    PrimaryQueryResult: PrimaryQueryResult;
}

export interface PrimaryQueryResult{
    RelevantResults: RelevantResults;
}

export interface RelevantResults {
    Table: Table;
    TotalRows: number;
}

export interface Table {
    Rows: Row[];
}

export interface ISPDate {
    value: string;
    error: Error;
}

export interface ISPUsers {
    value: ISPUser[];
    error: Error;
    message: string;
    OwnerTitle: string;
}

export interface Error{
    message: string;
    code: string;
}
export interface ISPUser {
    Title: string;
}
// export interface Rows {
//     value: Row[];
// }

export interface Row {
    Cells: Cell[];
    length: number;
}

export class State {
    private RootSite: ISPSite;
    private sites: ISPSite[];
    private searchResults: ISPSite[];
    private rowsPerPage: number = 50;

    constructor(site: Promise<ISPSite>) {
        this.RootSite = {
            DocID: "0",
            Title : "local",
            Url : "local",
            ViewsLifeTime : 3,
            ViewsRecent : 2,
            Size : 24,
            SiteDescription : "this is a local site and can not be determined",
            LastItemUserModifiedDate : new Date(),
            renderTemplateId : "templateid"
          };
        site.then((response) => {
            this.RootSite = response;
        });
        this.sites = [];
        this.searchResults = [];
    }



    public  addSite(site :ISPSite): void{
        
        this.sites.push(site);
        // this.sites.push(site);
    }

    public setSites(sites : ISPSite[]): void {
        this.sites = sites;
    }

    public sortSitesByLastItemUserModified(): void {
        this.searchResults = [];
        this.sites.sort((a: ISPSite, b: ISPSite) => {
            if(null == a.LastItemUserModifiedDate && null == b.LastItemUserModifiedDate){
                return 0 ;
            }
            else if(null == a.LastItemUserModifiedDate){
                return -b.LastItemUserModifiedDate.getTime();
            }
            else if(null == b.LastItemUserModifiedDate){
                return a.LastItemUserModifiedDate.getTime();
            }
            else{
                return a.LastItemUserModifiedDate.getTime() - b.LastItemUserModifiedDate.getTime();
            }
        });
        for(var i = 0; i < this.rowsPerPage; i++){
            this.searchResults.push(this.sites[i]);
        }

    }

    public sortSitesByRecentViews(): void {
        this.searchResults= [];
        this.sites.sort((a: ISPSite, b: ISPSite) => {
            return a.ViewsRecent - b.ViewsRecent;
        });
        for(var i = 0; i < this.rowsPerPage; i++){
            this.searchResults.push(this.sites[i]);
        }
    }

    public getRootSite(): ISPSite{
        return this.RootSite;
    }

    public getSites(): ISPSite[]{
        return this.sites;
    }

    public getSearchResults(): ISPSite[]{
        return this.searchResults;
    }
    public convertSPDate(date: string): Date {
        
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
}

export interface ISPSite{
    DocID: string;
    Title: string;
    Url: string;
    ViewsLifeTime: number;
    ViewsRecent: number;
    Size: number;
    SiteDescription: string;
    LastItemUserModifiedDate: Date;
    renderTemplateId: string;
}
export interface wanted{
    [Key: string] : boolean;
}


export interface Cell {
    Key: string;
    Value: string;
}

export interface Usage {
    Storage: number;
}

export interface SPGroupList {
    value: SPGroup[];
    error: Error;
    
    
}

export interface SPGroup  {
    Id: number;
    OwnerTitle: string;
}
