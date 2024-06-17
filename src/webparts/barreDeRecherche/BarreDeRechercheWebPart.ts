import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle, PropertyFieldMultiSelect } from "@pnp/spfx-property-controls";
import { PropertyPaneTextField, IPropertyPaneConfiguration, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPOperations } from '../Services/SPBarreRechercheService';
import { IDropdownOption } from "office-ui-fabric-react";

import * as strings from 'BarreDeRechercheWebPartStrings';
import BarreDeRecherche from './components/BarreDeRecherche';
import { IBarreDeRechercheProps } from './components/IBarreDeRechercheProps';

export interface IBarreDeRechercheWebPartProps {
  searchLabel: string;
  searchLabelColor: string;
  searchBarBackground: string;
  borderRadiusStyle: string;
  selectedLibraries: string[];
}

export default class BarreDeRechercheWebPart extends BaseClientSideWebPart<IBarreDeRechercheWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _spOperations: SPOperations;
  private _libraryOptions: IDropdownOption[] = [];

  constructor() {
    super();
    this._spOperations = new SPOperations();
  }


  public render(): void {
    const element: React.ReactElement<IBarreDeRechercheProps> = React.createElement(
      BarreDeRecherche,
      {
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        context: this.context,
        searchLabel: this.properties.searchLabel,
        searchLabelColor: this.properties.searchLabelColor,
        searchBarBackground: this.properties.searchBarBackground,
        borderRadiusStyle: this.properties.borderRadiusStyle,
        selectedLibraries: this.properties.selectedLibraries || [],
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      console.log(message);
      return this._loadLibraries();
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
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
                PropertyPaneTextField('searchLabel', {
                  label: "Search Label",
                  value: this.properties.searchLabel
                }),
                PropertyFieldColorPicker('searchLabelColor', {
                  label: "Search Label Color",
                  selectedColor: this.properties.searchLabelColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: 'searchLabelColorFieldId'
                }),
                PropertyFieldColorPicker('searchBarBackground', {
                  label: "Search Bar Background",
                  selectedColor: this.properties.searchBarBackground,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: 'searchBarBackgroundFieldId'
                }),
                PropertyFieldMultiSelect('selectedLibraries', {
                  label: "Select Libraries",
                  options: this._libraryOptions,
                  selectedKeys: this.properties.selectedLibraries,
                  key: 'selectedLibrariesFieldId'  // Add the key property
                }),
                PropertyPaneDropdown('borderRadiusStyle', {
                  label: "Search Bar Border Radius",
                  options: [
                    { key: '25px', text: 'Rounded' },
                    { key: '10px', text: 'Semi-Rounded' },
                    { key: '0px', text: 'Strict' }
                  ],
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  private _loadLibraries(): Promise<void> {
    return this._spOperations.GetAllList(this.context).then((libraries: IDropdownOption[]) => {
      this._libraryOptions = libraries;
      this.context.propertyPane.refresh();
    });
  }
}
