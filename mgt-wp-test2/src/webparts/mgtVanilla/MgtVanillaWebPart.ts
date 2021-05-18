import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'MgtVanillaWebPartStrings';
import styles from './MgtVanillaWebPart.module.scss';

export interface IMgtVanillaWebPartProps {
  description: string;
}

export default class MgtVanillaWebPart extends BaseClientSideWebPart<IMgtVanillaWebPartProps> {
  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.mgtVanilla}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Vanila webpart</span>
              <mgt-person person-query="me" show-name show-email></mgt-person>
            </div>
          </div>
        </div>
      </div>`;
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
