import { IDataProvider } from '../dataproviders/IDataProvider';
import { SPHttpClient } from '@microsoft/sp-http';
import { ISPSite } from '../common/IObjects';


export interface IManageListsProps {
  provider: IDataProvider;
  site: ISPSite;
  spHttpClient: SPHttpClient;

}
