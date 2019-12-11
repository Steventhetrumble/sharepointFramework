import { ISPList } from './GetListItemsWebPart';

export default class MockHttpClient {
    private static _items: ISPList[] = [{ 'EmployeeId': 'E123', 'EmployeeName': 'John', 'Experience': 'SharePoint', 'Location': 'India'},];
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {
        return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}

