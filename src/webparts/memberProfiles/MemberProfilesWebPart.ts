import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { MemberProfiles } from './components/MemberProfiles';
import type { IMemberProfilesProps } from './models';

export interface IMemberProfilesWebPartProps {
  listId: string;
  imageLibraryId?: string;
  itemsPerPage: number;
  accentColor: string;
  pageTitle?: string;
  subTitle?: string;
  showInMobile?: boolean;
}

export default class MemberProfilesWebPart
  extends BaseClientSideWebPart<IMemberProfilesWebPartProps> {

  /** Cached dropdown options */
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _imageListOptions: IPropertyPaneDropdownOption[] = [];
  private _listsLoaded = false;

  public async onInit(): Promise<void> {
    await this._ensureLists();
    return super.onInit();
  }

  public render(): void {
    const element = React.createElement(MemberProfiles, {
      listId: this.properties.listId,
      itemsPerPage: this.properties.itemsPerPage,
      accentColor: this.properties.accentColor,
      pageTitle: this.properties.pageTitle,
      subTitle: this.properties.subTitle,
      imageLibraryId: this.properties.imageLibraryId, 
      context: this.context
    } as IMemberProfilesProps);

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** ---------- Property Pane ---------- */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Select Member Profiles Data List' },
          groups: [
            {
              groupName: 'Source',
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: 'List',
                  options: this._listOptions,
                  disabled: this._listOptions.length === 0,
                }),
                PropertyPaneDropdown('imageLibraryId', {
                  label: 'Image library (optional)',
                  options: this._imageListOptions,
                  disabled: this._imageListOptions.length === 0,
                }),
              ],
            },
            {
              groupName: 'Display',
              groupFields: [
                PropertyPaneSlider('itemsPerPage', {
                  label: 'Max Users on a page (0 = All)',
                  min: 0, max: 2000, step: 1,
                }),
                PropertyPaneTextField('accentColor', {
                  label: 'Accent color (hex)',
                  description: 'Corporate color, e.g. #114461',
                }),
              ],
            },
            {
              groupName: 'Page Title',
              groupFields: [
                PropertyPaneTextField('pageTitle', { label: 'Page Title' }),
                PropertyPaneTextField('subTitle', { label: 'Sub title' }),
              ],
            },
            {
              groupName: 'Visibility',
              groupFields: [
                PropertyPaneToggle('showInMobile', {
                  label: 'Show in mobile and email view',
                  onText: 'On',
                  offText: 'Off',
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  /** Refresh lists when the pane is opened */
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this._ensureLists();
    this.context.propertyPane.refresh();
  }

  /** ---------- Helpers ---------- */

  /** Load both standard lists (BaseTemplate 100) and image/document libs (101/109) */
  private async _ensureLists(): Promise<void> {
    if (this._listsLoaded) return;

    try {
      const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');

      // Generic lists (100)
      const listsUrl =
        `${webUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,Hidden&$filter=Hidden eq false and BaseTemplate eq 100`;
      const listRes: SPHttpClientResponse = await this.context.spHttpClient.get(
        listsUrl, SPHttpClient.configurations.v1);
      const listJson: any = await listRes.json();
      const listOpts: IPropertyPaneDropdownOption[] = [];
      const listValues: any[] = listJson && listJson.value ? listJson.value : [];
      for (let i = 0; i < listValues.length; i++) {
        const r = listValues[i];
        listOpts.push({ key: r.Id, text: r.Title });
      }
      this._listOptions = listOpts;

      // Libraries: document (101) or picture (109)
      const libsUrl =
        `${webUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,Hidden&$filter=Hidden eq false and (BaseTemplate eq 101 or BaseTemplate eq 109)`;
      const libRes: SPHttpClientResponse = await this.context.spHttpClient.get(
        libsUrl, SPHttpClient.configurations.v1);
      const libJson: any = await libRes.json();
      const libOpts: IPropertyPaneDropdownOption[] = [];
      const libValues: any[] = libJson && libJson.value ? libJson.value : [];
      for (let i = 0; i < libValues.length; i++) {
        const r = libValues[i];
        libOpts.push({ key: r.Id, text: r.Title });
      }
      this._imageListOptions = libOpts;

      // Preselect first values if none chosen
      if (!this.properties.listId && this._listOptions.length) {
        this.properties.listId = String(this._listOptions[0].key);
      }
      if (!this.properties.itemsPerPage && this.properties.itemsPerPage !== 0) {
        this.properties.itemsPerPage = 0; // show all by default
      }
      if (!this.properties.accentColor) {
        this.properties.accentColor = '#114461';
      }

      this._listsLoaded = true;
    } catch (e) {
      // Swallow errors so the pane still renders; options will be empty
      // eslint-disable-next-line no-console
      console.warn('Failed to load lists/libraries', e);
      this._listOptions = [];
      this._imageListOptions = [];
      this._listsLoaded = true;
    }
  }
}
