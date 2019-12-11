import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ManageListsWebPartStrings';
import ManageLists from './components/ManageLists';
import { IManageListsProps } from './components/IManageListsProps';
import MockupDataProvider from './dataproviders/MockupDataProvider';
import { SPHttpClient } from '@microsoft/sp-http';



export interface IManageListsWebPartProps {
  description: string;
}

export default class ManageListsWebPart extends BaseClientSideWebPart<IManageListsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IManageListsProps > = React.createElement(
      ManageLists,
      {
        provider: new MockupDataProvider(),
        site: {
          DocID: this.context.pageContext.web.id.toString(),
          Title: this.context.pageContext.web.title,
          Url: this.context.pageContext.web.absoluteUrl,
          ViewsLifeTime: 0,
          ViewsRecent: 0,
          Size: 0,
          SiteDescription: this.context.pageContext.web.description,
          LastItemUserModifiedDateSharepoint: null,
          LastItemUserModifiedDate: null,
          LastItemUserModifiedDateFomatted: null,
          LastItemUserModifiedDatevalue: null,
          renderTemplateId: this.context.pageContext.web.templateName
        },
        spHttpClient: this.context.spHttpClient

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
