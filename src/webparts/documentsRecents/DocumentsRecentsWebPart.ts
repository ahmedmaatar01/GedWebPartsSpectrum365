import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IDropdownOption } from "office-ui-fabric-react";
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from "@pnp/spfx-property-controls";

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPOperations } from '../Services/SPDocRecentsServices';

import * as strings from 'DocumentsRecentsWebPartStrings';
import DocumentsRecents from './components/DocumentsRecents';
import { IDocumentsRecentsProps } from './components/IDocumentsRecentsProps';

export interface IDocumentsRecentsWebPartProps {
  description: string;
  list_titles: IDropdownOption[];
  list_title: string;
  backgroundColor: string;
  textColor: string;
  timeInterval: string; // New property
  numFields: number; // New property
}

export default class DocumentsRecentsWebPart extends BaseClientSideWebPart<IDocumentsRecentsWebPartProps> {
  private _spOperations: SPOperations;
  private _validateNumberField(value: string): string {
    const num = parseInt(value, 10);
    if (isNaN(num) || num < 1 || num > 50) {
      return 'Please enter a number between 1 and 50.';
    }
    return '';
  }

  constructor() {
    super();
    this._spOperations = new SPOperations();
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IDocumentsRecentsProps> = React.createElement(
      DocumentsRecents,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        list_title: this.properties.list_title,
        backgroundColor: this.properties.backgroundColor,
        textColor: this.properties.textColor,
        timeInterval: this.properties.timeInterval, // Pass new property
        numFields: this.properties.numFields // Pass new property
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;

      // Set default background color if not set
      if (!this.properties.backgroundColor) {
        this.properties.backgroundColor = '#3c3b5e';
      }
      if (!this.properties.textColor) {
        this.properties.textColor = '#ffffff';
      }

      // Fetch all lists
      return this._spOperations.GetAllList(this.context)
        .then((result: IDropdownOption[]) => {
          this.properties.list_titles = result;
        });
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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
    const { semanticColors } = currentTheme;

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
                PropertyPaneDropdown('list_title', {
                  label: "Select a title",
                  options: this.properties.list_titles,
                  selectedKey: this.properties.list_title,
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                
                PropertyPaneDropdown('timeInterval', {
                  label: "Select time interval",
                  options: [
                    { key: 'today', text: 'Today' },
                    { key: 'lastWeek', text: 'Last week' },
                    { key: 'last15Days', text: 'Last 15 days' },
                    { key: 'lastMonth', text: 'Last month' },
                    { key: 'last3Months', text: 'Last 3 months' },
                    { key: 'last6Months', text: 'Last 6 months' },
                    { key: 'lastYear', text: 'Last year' },
                  ],
                  selectedKey: 'lastWeek'
                }),
                PropertyPaneTextField('numFields', {
                  label: "Number of fields to display",
                  value: '3',
                  onGetErrorMessage: this._validateNumberField,
                  deferredValidationTime: 500
                }),
                PropertyFieldColorPicker('backgroundColor', {
                  label: "Select background color",
                  selectedColor: this.properties.backgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: 'backgroundColorFieldId'
                }),
                PropertyFieldColorPicker('textColor', {
                  label: "Select text color",
                  selectedColor: this.properties.textColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: 'textColorFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
  

  

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'backgroundColor' && newValue !== oldValue) {
      this.properties.backgroundColor = newValue;
      this.render();
    } else if (propertyPath === 'textColor' && newValue !== oldValue) {
      this.properties.textColor = newValue;
      this.render();
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}
