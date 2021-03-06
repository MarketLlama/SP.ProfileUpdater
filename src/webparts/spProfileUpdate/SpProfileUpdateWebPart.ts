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
import { loadTheme } from 'office-ui-fabric-react';
import { initializeIcons } from '@uifabric/icons';
initializeIcons();

loadTheme({
  palette: {
    themePrimary: '#5f7800',
    themeLighterAlt: '#f7faf0',
    themeLighter: '#e1e9c4',
    themeLight: '#c9d696',
    themeTertiary: '#97ae46',
    themeSecondary: '#6e8810',
    themeDarkAlt: '#546c00',
    themeDark: '#475b00',
    themeDarker: '#354300',
    neutralLighterAlt: '#f8f8f8',
    neutralLighter: '#f4f4f4',
    neutralLight: '#eaeaea',
    neutralQuaternaryAlt: '#dadada',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c8c8',
    neutralTertiary: '#c2c2c2',
    neutralSecondary: '#858585',
    neutralPrimaryAlt: '#4b4b4b',
    neutralPrimary: '#333333',
    neutralDark: '#272727',
    black: '#1d1d1d',
    white: '#ffffff',
  }
});

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
