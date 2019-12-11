import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxgraphclientSampleWebPartStrings';
import SpfxgraphclientSample from './components/SpfxgraphclientSample';
import { ISpfxgraphclientSampleProps } from './components/ISpfxgraphclientSampleProps';
import { ClientMode } from './components/ClientMode';

export interface ISpfxgraphclientSampleWebPartProps {
  clientMode: ClientMode;
  description: string;
}

export default class SpfxgraphclientSampleWebPart extends BaseClientSideWebPart<ISpfxgraphclientSampleWebPartProps> {

  public render(): void {
   
    const element: React.ReactElement<ISpfxgraphclientSampleProps > = React.createElement(
      SpfxgraphclientSample,
      {
        description: this.properties.description,
        clientMode: this.properties.clientMode,
        context: this.context
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
                PropertyPaneChoiceGroup('clientMode', {
                  label: strings.ClientModeLabel,
                  options: [
                    {key: ClientMode.aad, text: "AadHttpClient"},
                    { key: ClientMode.graph, text: "MSGraphclient"}
                  ]
                }),
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
