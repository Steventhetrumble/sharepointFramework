import {
    ISPSite
} from '../../common/IObjects';
import { SPHttpClient } from '@microsoft/sp-http';

export default interface ISPTitleBlockProps {
    site: ISPSite;
    spHttpClient: SPHttpClient;
}