import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import MemberProfiles from './components/MemberProfiles';
import type { IMemberProfilesProps } from './components/IMemberProfilesProps';
import { SpService } from './services/SpService';

export interface IMemberProfilesWebPartProps extends IMemberProfilesProps {}

export default class MemberProfilesWebPart extends BaseClientSideWebPart<IMemberProfilesWebPartProps> {
  private listOptions: IPropertyPaneDropdownOption[] = [];
  private imageListOptions: IPropertyPaneDropdownOption[] = [];
  private svc!: SpService;

  public async onInit(): Promise<void> {
    this.svc = new SpService(this.context);
  }

  public render(): void {
    const element = React.createElement(MemberProfiles as any, {
      context: this.context,
      listId: this.properties.listId,
      imageListId: this.properties.imageListId,                // may be undefined
      itemsPerPage: this.properties.itemsPerPage,              // 0/undefined => show all
      accentColor: this.properties.accentColor || '#114461'
    } as IMemberProfilesProps);

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if (this.listOptions.length === 0) {
      this.listOptions = await this.svc.getLists();
    }
    if (this.imageListOptions.length === 0) {
      const imgs = await this.svc.getImageLibraries();
      this.imageListOptions = [{ key: '', text: '(None)' }].concat(imgs);
    }
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Select Member Profiles Data List' },
          groups: [
            {
              groupName: 'Source',
              groupFields: [
                PropertyPaneDropdown('listId', { label: 'List', options: this.listOptions }),
                PropertyPaneDropdown('imageListId', { label: 'Image library (optional)', options: this.imageListOptions }),
              ]
            },
            {
              groupName: 'Display',
              groupFields: [
                // allow large numbers; 0 === ALL
                PropertyPaneSlider('itemsPerPage', { label: 'Max Users on a page (0 = All)', min: 0, max: 5000, step: 10 }),
                PropertyPaneTextField('accentColor', { label: 'Accent color (hex)', description: 'Corporate color, e.g. #114461' })
              ]
            }
          ]
        }
      ]
    };
  }
}
