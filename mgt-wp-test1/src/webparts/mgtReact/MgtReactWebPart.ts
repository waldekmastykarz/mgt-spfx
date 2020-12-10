import { Providers, SharePointProvider } from 'mgt-spfx';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'MgtReactWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IMgtReactProps } from './components/IMgtReactProps';
import MgtReact from './components/MgtReact';

export interface IMgtReactWebPartProps {
  description: string;
}

export default class MgtReactWebPart extends BaseClientSideWebPart<IMgtReactWebPartProps> {
  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context);
  }

  public render(): void {
    const element: React.ReactElement<IMgtReactProps> = React.createElement(
      MgtReact,
      {
        description: this.properties.description
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
