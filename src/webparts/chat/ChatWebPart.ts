import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ThemeProvider, type ThemeChangedEventArgs, type IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ChatWebPartStrings';
import Chat from './components/Chat';
import { type IChatProps, type ChatTheme } from './components/Chat';
import { SetupService } from '../../services/SetupService';
import { GraphService } from '../../services/GraphService';

export interface IChatWebPartProps {
  useSiteTheme?: boolean;
  buttonPrimaryColor?: string;
  buttonTextColor?: string;
  surfaceBorderColor?: string;
  messageBorderColor?: string;
  selfMessageBackgroundColor?: string;
  mentionBackgroundColor?: string;
  mentionTextColor?: string;
}

export default class ChatWebPart extends BaseClientSideWebPart<IChatWebPartProps> {

  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();

    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChanged);

    SetupService.init(this.context);
    await SetupService.ensureList();
    await GraphService.init(this.context);
  }
  
  public render(): void {
    const palette = this.properties.useSiteTheme && this._themeVariant ? this._themeVariant.palette : undefined;

    const theme: ChatTheme = {
      primaryButtonBackground: this.properties.buttonPrimaryColor || palette?.themePrimary || '#0078d4',
      primaryButtonText: this.properties.buttonTextColor || palette?.white || '#ffffff',
      surfaceBorderColor: this.properties.surfaceBorderColor || palette?.neutralLight || '#dddddd',
      messageBorderColor: this.properties.messageBorderColor || palette?.neutralLighter || '#eeeeee',
      selfMessageBackground: this.properties.selfMessageBackgroundColor || palette?.themeLighterAlt || '#f3f2f1',
      mentionBackground: this.properties.mentionBackgroundColor || palette?.themeLighter || '#e8f3ff',
      mentionText: this.properties.mentionTextColor || palette?.themeDarker || '#004578'
    };

    const element: React.ReactElement<IChatProps> = React.createElement(Chat, {
        context: this.context,
        theme
      } as IChatProps);

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    if (this._themeProvider) {
      this._themeProvider.themeChangedEvent.remove(this, this._handleThemeChanged);
    }
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
                PropertyPaneToggle('useSiteTheme', {
                  label: strings.UseSiteThemeLabel
                }),
                PropertyPaneTextField('buttonPrimaryColor', {
                  label: strings.PrimaryButtonColorLabel,
                  description: strings.ColorHelpText
                }),
                PropertyPaneTextField('buttonTextColor', {
                  label: strings.PrimaryButtonTextColorLabel,
                  description: strings.ColorHelpText
                }),
                PropertyPaneTextField('surfaceBorderColor', {
                  label: strings.SurfaceBorderColorLabel,
                  description: strings.ColorHelpText
                }),
                PropertyPaneTextField('messageBorderColor', {
                  label: strings.MessageBorderColorLabel,
                  description: strings.ColorHelpText
                }),
                PropertyPaneTextField('selfMessageBackgroundColor', {
                  label: strings.SelfMessageBackgroundColorLabel,
                  description: strings.ColorHelpText
                }),
                PropertyPaneTextField('mentionBackgroundColor', {
                  label: strings.MentionBackgroundColorLabel,
                  description: strings.ColorHelpText
                }),
                PropertyPaneTextField('mentionTextColor', {
                  label: strings.MentionTextColorLabel,
                  description: strings.ColorHelpText
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _handleThemeChanged(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }
}
