export interface IList{
    Title?: string;
    Id?: string;
}

export interface ISPUsers {
    value: ISPUser[];
    error: Error;
    message: string;
    OwnerTitle: string;
}
export interface ISPUser {
    Title: string;
    Email: string;
}
export interface ISPDate {
    date: Date;
    value: string;
    dateFormatted:string;
    datevalue:number;
    error: Error;
}

export interface ISPGroupEmails{
    groups: ISPGroupEmail[];
}

export interface ISPGroupEmail {
    GroupName: string;
    userEmails: string[];
}

// export class searchState {
//     private RootSite: ISPSite;
//     private sites: ISPSite[];
//     private searchResults: ISPSite[];
//     private rowsPerPage: number = 50;

//     constructor(site: Promise<ISPSite>) {
//         this.RootSite = {
//             DocID: "0",
//             Title : "local",
//             Url : "local",
//             ViewsLifeTime : 3,
//             ViewsRecent : 2,
//             Size : 24,
//             SiteDescription : "this is a local site and can not be determined",
//             LastItemUserModifiedDate : new Date(),
//             renderTemplateId : "templateid"
//           };
//         site.then((response) => {
//             this.RootSite = response;
//         });
//         this.sites = [];
//         this.searchResults = [];
//     }



//     public  addSite(site :ISPSite): void{
        
//         this.sites.push(site);
//         // this.sites.push(site);
//     }

//     public setSites(sites : ISPSite[]): void {
//         this.sites = sites;
//     }

//     public sortSitesByLastItemUserModified(): void {
//         this.searchResults = [];
//         this.sites.sort((a: ISPSite, b: ISPSite) => {
//             if(null == a.LastItemUserModifiedDate && null == b.LastItemUserModifiedDate){
//                 return 0 ;
//             }
//             else if(null == a.LastItemUserModifiedDate){
//                 return -b.LastItemUserModifiedDate.getTime();
//             }
//             else if(null == b.LastItemUserModifiedDate){
//                 return a.LastItemUserModifiedDate.getTime();
//             }
//             else{
//                 return a.LastItemUserModifiedDate.getTime() - b.LastItemUserModifiedDate.getTime();
//             }
//         });
//         for(var i = 0; i < this.rowsPerPage; i++){
//             this.searchResults.push(this.sites[i]);
//         }

//     }

//     public sortSitesByRecentViews(): void {
//         this.searchResults= [];
//         this.sites.sort((a: ISPSite, b: ISPSite) => {
//             return a.ViewsRecent - b.ViewsRecent;
//         });
//         for(var i = 0; i < this.rowsPerPage; i++){
//             this.searchResults.push(this.sites[i]);
//         }
//     }

//     public getRootSite(): ISPSite{
//         return this.RootSite;
//     }

//     public getSites(): ISPSite[]{
//         return this.sites;
//     }

//     public getSearchResults(): ISPSite[]{
//         return this.searchResults;
//     }
//    
// }

export interface ISPSite{
    DocID: string;
    Title: string;
    Url: string;
    ViewsLifeTime: number;
    ViewsRecent: number;
    Size: number;
    SiteDescription: string;
    LastItemUserModifiedDateSharepoint: string;
    LastItemUserModifiedDateFomatted: string;
    LastItemUserModifiedDatevalue: number;
    LastItemUserModifiedDate: Date;
    renderTemplateId: string;
}

export interface SPGroupList {
    value: SPGroup[];
    error: Error;
    
    
}

export interface SPGroup  {
    Id: number;
    OwnerTitle: string;
    users: ISPUsers;
}

export interface Usage {
    Storage: number;
}