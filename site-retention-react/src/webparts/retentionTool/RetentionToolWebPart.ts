import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from "@microsoft/sp-webpart-base";
import * as strings from "RetentionToolWebPartStrings";
import RetentionTool from "./components/RetentionTool";
import { IRetentionToolProps } from "./components/IRetentionTool.types";

export interface IRetentionToolWebPartProps {
  description: string;
}

export default class RetentionToolWebPart extends BaseClientSideWebPart<
  IRetentionToolWebPartProps
> {
  public render(): void {
    const element: React.ReactElement<
      IRetentionToolProps
    > = React.createElement(RetentionTool, {
      description: this.properties.description,
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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
