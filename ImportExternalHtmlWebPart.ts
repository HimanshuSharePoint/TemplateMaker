import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import Objdashboardtemplate from './dashboard-template';

import * as strings from 'ImportExternalHtmlWebPartStrings';

export interface IImportExternalHtmlWebPartProps {
  description: string;
}

export default class ImportExternalHtmlWebPart extends BaseClientSideWebPart<IImportExternalHtmlWebPartProps> {

    public render(): void {
        this.domElement.innerHTML = Objdashboardtemplate.dashboardHTML;
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
