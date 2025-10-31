import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ChatWebPartStrings';
import Chat from './components/Chat';
import { IChatProps } from './components/Chat';
import { SetupService } from '../../services/SetupService';
import { GraphService } from '../../services/GraphService';


export default class ChatWebPart extends BaseClientSideWebPart<{}> {

  protected async onInit(): Promise<void> {
    await super.onInit();
    SetupService.init(this.context);
    await SetupService.ensureList();
    await GraphService.init(this.context);
  }
  
  public render(): void {
    const element: React.ReactElement<IChatProps> = React.createElement(Chat, {
        context: this.context
      } as IChatProps);

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
