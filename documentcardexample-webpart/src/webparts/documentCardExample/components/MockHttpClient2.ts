import { ISPSite } from './GetListItemsWebPart';

export default class MockHttpClient2 {
    private static _items: ISPSite = {
        DocID : "0",
        Title : "local",
        Url : "local",
        ViewsLifeTime : 3,
        ViewsRecent : 2,
        Size : 24,
        SiteDescription : "this is a local site and can not be determined",
        LastItemUserModifiedDate : new Date(),
        renderTemplateId : "templateid"
    };
    public static get(restUrl: string, options?: any): Promise<ISPSite> {
        return new Promise<ISPSite>((resolve) => {
            resolve(MockHttpClient2._items);
        });
    }
}