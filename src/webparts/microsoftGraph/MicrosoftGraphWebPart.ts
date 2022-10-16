import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MicrosoftGraphWebPartStrings';
import MicrosoftGraph from './components/MicrosoftGraph';
import { IMicrosoftGraphProps } from './components/IMicrosoftGraphProps';
import { ClientMode } from './components/ClientMode';

export interface IMicrosoftGraphWebPartProps {
  clientMode: ClientMode;
}

export default class MicrosoftGraphWebPart extends BaseClientSideWebPart<IMicrosoftGraphWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMicrosoftGraphProps> = React.createElement(
      MicrosoftGraph,
      {
        clientMode: this.properties.clientMode,
        context: this.context,
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
                    { key: ClientMode.aad, text: "AadHttpClient"},
                    { key: ClientMode.graph, text: "MSGraphClient"},
                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
