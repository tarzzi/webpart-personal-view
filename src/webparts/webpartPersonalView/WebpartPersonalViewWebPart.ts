import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WebpartPersonalViewWebPartStrings';
import WebpartPersonalView from './components/WebpartPersonalView';
import { IWebpartPersonalViewProps } from './components/IWebpartPersonalViewProps';

export interface IWebpartPersonalViewWebPartProps {
  description: string;
  greetingPrefix: string;
  greetingSuffix: string;
  greetingShowUser: boolean;
  subGreeting: string;
  showGreeting: boolean;
  mailRetrieveCount: number;
  eventRetrieveCount: number;
  fileRetrieveCount: number;
  mailCount: number;
  eventCount: number;
  fileCount: number;
}

export default class WebpartPersonalViewWebPart extends BaseClientSideWebPart<IWebpartPersonalViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IWebpartPersonalViewProps> = React.createElement(
      WebpartPersonalView,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        greetingPrefix: this.properties.greetingPrefix,
        greetingSuffix: this.properties.greetingSuffix,
        greetingShowUser: this.properties.greetingShowUser,
        subGreeting: this.properties.subGreeting,
        showGreeting: this.properties.showGreeting,
        mailRetrieveCount: this.properties.mailCount,
        eventRetrieveCount: this.properties.eventCount,
        fileRetrieveCount: this.properties.fileCount
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
                PropertyPaneToggle('showGreeting', {
                  label: strings.ShowGreetingFieldLabel,
                  onText: 'Show',
                  offText: 'Hide',
                  checked: true
                }),
                PropertyPaneTextField('greetingPrefix', {
                  label: strings.GreetingFieldPrefix
                }),
                PropertyPaneTextField('greetingSuffix', {
                  label: strings.GreetingFieldSuffix
                }),
                PropertyPaneToggle('greetingShowUser', {
                  label: strings.GreetingFieldShowUser,
                  onText: 'Show',
                  offText: 'Hide',
                  checked: true
                }),
                PropertyPaneTextField('subGreeting', {
                  label: strings.SubGreetingFieldLabel
                })
              ]
            },
            {
              groupName: strings.RetrievableItemsGroupName,
              groupFields: [
                PropertyPaneSlider('mailCount', {
                  label: strings.MailRetrieveCountFieldLabel,
                  min: 1,
                  max: 20,
                  value: 5
                }),
                PropertyPaneSlider('eventCount', {
                  label: strings.EventRetrieveCountFieldLabel,
                  min: 1,
                  max: 20,
                  value: 10
                }),
                PropertyPaneSlider('fileCount', {
                  label: strings.FileRetrieveCountFieldLabel,
                  min: 1,
                  max: 20,
                  value: 6
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
