import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import pnp from '@pnp/pnpjs';
import * as strings from 'SpProfileUpdateWebPartStrings';
import SpProfileUpdate from './components/SpProfileUpdate';
import { ISpProfileUpdateProps } from './components/ISpProfileUpdateProps';

export interface ISpProfileUpdateWebPartProps {
  description: string;
}

export default class SpProfileUpdateWebPart extends BaseClientSideWebPart<ISpProfileUpdateWebPartProps> {

  public render(): void {
    let context = pnp.sp.site.getContextInfo();
    const element: React.ReactElement<ISpProfileUpdateProps > = React.createElement(
      SpProfileUpdate,
      {
        description: this.properties.description,
        context : this.context
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
