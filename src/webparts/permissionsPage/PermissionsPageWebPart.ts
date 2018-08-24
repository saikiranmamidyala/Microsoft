import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PermissionsPageWebPartStrings';
import PermissionsPage from './components/PermissionsPage';
import { IPermissionsPageProps } from './components/IPermissionsPageProps';
import { fixPageStyling } from '../../shared/util';
import Teams from '../../shared/Teams';
import pnp from 'sp-pnp-js/lib/pnp';
import { clientTag } from '../../shared/global';

export interface IPermissionsPageWebPartProps {
  description: string;
}

export default class PermissionsPageWebPart extends BaseClientSideWebPart<IPermissionsPageWebPartProps> {
  public onInit() {
    return Promise.resolve().then(() => {
      fixPageStyling()
      Teams.initialize()
      pnp.setup({
        spfxContext: this.context,
        globalCacheDisable: true,
        sp: {
          headers: {
            "X-ClientTag": clientTag,
            "User-Agent": clientTag
          }
        }
      })
    })
  }

  public render(): void {
    const element: React.ReactElement<IPermissionsPageProps > = React.createElement(
      PermissionsPage,
      {
        description: this.properties.description,
        siteTitle: this.context.pageContext.web.title,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion1(): Version {
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
