import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'RemoveTabPageWebPartStrings';
import RemoveTabPage from './components/RemoveTabPage';
import { IRemoveTabPageProps } from './components/IRemoveTabPageProps';
import { fixPageStyling } from '../../shared/util';
import Teams from '../../shared/Teams';
import pnp from 'sp-pnp-js/lib/pnp';
import { clientTag } from '../../shared/global';

export interface IRemoveTabPageWebPartProps {
  description: string;
}

export default class RemoveTabPageWebPart extends BaseClientSideWebPart<IRemoveTabPageWebPartProps> {
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
    const element: React.ReactElement<IRemoveTabPageProps > = React.createElement(
      RemoveTabPage,
      {
        description: this.properties.description
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
