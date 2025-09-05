import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
//import { BaseClientSideWebPart, PropertyPaneDropdown, IPropertyPaneDropdownOption, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { BaseClientSideWebPart, PropertyPaneDropdown, IPropertyPaneDropdownOption, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { MemberProfiles } from './components/MemberProfiles';
import type { IMemberProfilesProps } from './models';
import { SpService } from './services/SpService';

export interface IMemberProfilesWebPartProps extends IMemberProfilesProps {}

export default class MemberProfilesWebPart extends BaseClientSideWebPart<IMemberProfilesWebPartProps> {
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _loadingLists = false;

  public render(): void {
    const element = React.createElement(MemberProfiles, {
      listId: this.properties.listId,
      itemsPerPage: this.properties.itemsPerPage || 36,
      accentColor: this.properties.accentColor || '#114461',
      context: this.context
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const disabled = this._loadingLists;
    return {
      pages: [
        {
          header: { description: 'Select Member Profiles Data List' },
          groups: [
            {
              groupName: 'Source',
              groupFields: [
                PropertyPaneDropdown('listId', { label: 'List', options: this._listOptions, disabled })
              ]
            },
            {
              groupName: 'Display',
              groupFields: [
                //PropertyPaneSlider('itemsPerPage', { label: 'Max items', min: 12, max: 500, step: 10 }),
                PropertyPaneTextField('accentColor', { label: 'Accent color (hex)', description: 'Corporate color, e.g. #114461' })
              ]
            }
          ]
        }
      ]
    };
  }

  protected async loadPropertyPaneResources(): Promise<void> { await this._loadLists(); }
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if (this._listOptions.length === 0) { await this._loadLists(); this.context.propertyPane.refresh(); }
  }

  private async _loadLists(): Promise<void> {
    if (this._loadingLists) return; this._loadingLists = true;
    try { const svc = new SpService(this.context); this._listOptions = await svc.getLists(); }
    catch { this._listOptions = [{ key: '', text: 'Failed to load lists' }]; }
    finally { this._loadingLists = false; }
  }
}