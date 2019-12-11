import * as React from 'react';
import styles from '../ManageLists.module.scss';
import ISPTitleBlockProps from './ISPTitleBlockProps';
import ISPTitleBlockState from './ISPTitleBlockState';
import { IList, ISPSite, Usage } from '../../common/IObjects';
import {
    Version,
    Environment,
    EnvironmentType
} from '@microsoft/sp-core-library';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';

export default class SPTitleBlock extends React.Component<ISPTitleBlockProps, ISPTitleBlockState> {
    constructor(props: ISPTitleBlockProps) {
        super(props);
        this.state = {
            site: this.props.site   
        };
    }

    private _getRootSiteData(): Promise<ISPSite> {
        return this.props.spHttpClient.get(this.props.site.Url + '/_api/site/usage', SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
            return response.json().then((responseJSON: Usage) => {
              console.log(responseJSON);
              // TODO views lifetime and views recent 
              const site: ISPSite = {
                DocID: this.props.site.DocID,
                Title: this.props.site.Title,
                Url: this.props.site.Url,
                ViewsLifeTime: 0,
                ViewsRecent: 0,
                Size: responseJSON.Storage,
                SiteDescription: this.props.site.SiteDescription,
                LastItemUserModifiedDateSharepoint: null,
                LastItemUserModifiedDate: null,
                LastItemUserModifiedDateFomatted: null,
                LastItemUserModifiedDatevalue: null,
                renderTemplateId: this.props.site.renderTemplateId
              };
              return site;
            });
          });
      }

    public componentDidMount() {
        if (Environment.type === EnvironmentType.Local) {
            let tempSite: ISPSite = {
                DocID: "01",
                Title: "local",
                Url: "local",
                ViewsLifeTime: 31,
                ViewsRecent: 21,
                Size: 241,
                SiteDescription: "this is a local site and can not be described",
                LastItemUserModifiedDateSharepoint: '',
                LastItemUserModifiedDate: null,
                LastItemUserModifiedDatevalue: null,
                LastItemUserModifiedDateFomatted: "0",
                renderTemplateId: "templateid"
              };
            this.setState({
                site: tempSite
            });
        }
        else {
            this._getRootSiteData().then((result: ISPSite) => {
                this.setState({
                    site: result 
                });
            });
        }
    }


    public render(): React.ReactElement<ISPTitleBlockProps> {
        return (
            <div className={styles.titleBlock}>
                
                    <div className= {styles.row}>Title: {this.state.site.Title.toString()}</div>
                    <div className= {styles.row}>Url :{this.state.site.Url.toString()}</div>
                    <div className= {styles.row}>Size: {this.state.site.Size.toString()}</div>
                    <div className= {styles.row}>site description: {this.state.site.SiteDescription.toString()}</div>
                    <div className= {styles.row}>render template id: {this.state.site.renderTemplateId.toString()}</div>
            </div>
            );
    }
}
