import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CopilotStudioAgentWebPartStrings';
import CopilotStudioAgent from './components/CopilotStudioAgent';
import { ICopilotStudioAgentProps } from './components/ICopilotStudioAgentProps';
import { DEFAULT_APP_CLIENT_ID, DEFAULT_TENANT_ID } from '../../config/defaults';

export interface ICopilotStudioAgentWebPartProps {
  title: string;
  mode: string;
  connectionString: string;
  embedUrl: string;
  environmentId: string;
  agentIdentifier: string;
  appClientId: string;
  tenantId: string;
  height: number;
  useFluentTheme: boolean;
  showHeader: boolean;
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Accepts either the full embed code snippet (<iframe src="...">) or just the URL. */
function extractEmbedUrl(value: string): string {
  const v = (value || '').trim();
  const match = /src\s*=\s*["']([^"']+)["']/i.exec(v);
  return (match ? match[1] : v).replace(/&amp;/g, '&');
}

export default class CopilotStudioAgentWebPart extends BaseClientSideWebPart<ICopilotStudioAgentWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _accentColor: string | undefined;

  public render(): void {
    const element: React.ReactElement<ICopilotStudioAgentProps> = React.createElement(
      CopilotStudioAgent,
      {
        title: this.properties.title,
        mode: this.properties.mode === 'embed' ? 'embed' : 'sdk',
        connectionString: this.properties.connectionString,
        embedUrl: extractEmbedUrl(this.properties.embedUrl),
        environmentId: this.properties.environmentId,
        agentIdentifier: this.properties.agentIdentifier,
        appClientId: (this.properties.appClientId || DEFAULT_APP_CLIENT_ID || '').trim(),
        tenantId: (this.properties.tenantId || DEFAULT_TENANT_ID || this.context.pageContext.aadInfo.tenantId.toString()).trim(),
        height: this.properties.height,
        useFluentTheme: this.properties.useFluentTheme !== false,
        showHeader: this.properties.showHeader !== false,
        userLoginName: this.context.pageContext.user.loginName,
        userDisplayName: this.context.pageContext.user.displayName,
        locale: this.context.pageContext.cultureInfo.currentUICultureName || 'en-US',
        isEditMode: this.displayMode === 2,
        isDarkTheme: this._isDarkTheme,
        accentColor: this._accentColor,
        onConfigure: () => this.context.propertyPane.open()
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    this._accentColor = currentTheme.palette ? currentTheme.palette.themePrimary : undefined;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Apply property changes only when the author selects "Apply" (avoids reconnecting on every keystroke).
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  private _validateConnectionString(value: string): string {
    const v = (value || '').trim();
    if (!v) { return ''; }
    return /^https:\/\/[^\s]+\.api\.powerplatform\.com\/copilotstudio\//i.test(v) ? '' : strings.ConnectionStringInvalid;
  }

  private _validateGuid(value: string): string {
    const v = (value || '').trim();
    return !v || GUID.test(v) ? '' : strings.GuidInvalid;
  }

  private _validateEmbed(value: string): string {
    const url = extractEmbedUrl(value);
    if (!url) { return ''; }
    return /^https:\/\/([a-z0-9-]+\.)*(copilotstudio(\.preview)?\.microsoft\.com|powerva\.microsoft\.com)\//i.test(url) ? '' : strings.EmbedUrlInvalid;
  }

  private _agentFields(): IPropertyPaneField<unknown>[] {
    const fields: IPropertyPaneField<unknown>[] = [
      PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
      PropertyPaneDropdown('mode', {
        label: strings.ModeFieldLabel,
        options: [
          { key: 'sdk', text: strings.ModeSdk },
          { key: 'embed', text: strings.ModeEmbed }
        ]
      })
    ];
    if (this.properties.mode === 'embed') {
      fields.push(PropertyPaneTextField('embedUrl', {
        label: strings.EmbedUrlFieldLabel,
        description: strings.EmbedUrlFieldDescription,
        multiline: true,
        rows: 5,
        onGetErrorMessage: this._validateEmbed.bind(this),
        deferredValidationTime: 500
      }));
      return fields;
    }
    fields.push(
      PropertyPaneTextField('connectionString', {
        label: strings.ConnectionStringFieldLabel,
        description: strings.ConnectionStringFieldDescription,
        multiline: true,
        rows: 5,
        onGetErrorMessage: this._validateConnectionString.bind(this),
        deferredValidationTime: 500
      }),
      PropertyPaneLabel('orLabel', { text: strings.OrLabel }),
      PropertyPaneTextField('environmentId', {
        label: strings.EnvironmentIdFieldLabel,
        onGetErrorMessage: this._validateGuid.bind(this)
      }),
      PropertyPaneTextField('agentIdentifier', { label: strings.AgentIdentifierFieldLabel })
    );
    return fields;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const groups = [
      { groupName: strings.AgentGroupName, groupFields: this._agentFields() }
    ];
    if (this.properties.mode !== 'embed') {
      groups.push({
        groupName: strings.AuthGroupName,
        groupFields: [
          PropertyPaneTextField('appClientId', {
            label: strings.AppClientIdFieldLabel,
            description: strings.AppClientIdFieldDescription,
            placeholder: DEFAULT_APP_CLIENT_ID,
            onGetErrorMessage: this._validateGuid.bind(this)
          }),
          PropertyPaneTextField('tenantId', {
            label: strings.TenantIdFieldLabel,
            description: strings.TenantIdFieldDescription,
            placeholder: this.context.pageContext.aadInfo.tenantId.toString()
          })
        ]
      });
    }
    groups.push({
      groupName: strings.AppearanceGroupName,
      groupFields: [
        PropertyPaneToggle('showHeader', { label: strings.ShowHeaderFieldLabel }),
        PropertyPaneToggle('useFluentTheme', { label: strings.FluentThemeFieldLabel }),
        PropertyPaneSlider('height', { label: strings.HeightFieldLabel, min: 320, max: 1200, step: 20 })
      ]
    });
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups
        }
      ]
    };
  }
}
