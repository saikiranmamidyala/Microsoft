import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Version } from '@microsoft/sp-core-library'
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base'

import ExecutiveAssetsPage, {
  IExecutiveAssetsPageProps
} from './components/ExecutiveAssetsPage'

//import ExecutiveAssetsPage from './components/ExecutiveAssetsPage'

import Teams from '../../shared/Teams'
import { fixPageStyling } from '../../shared/util'
import pnp from 'sp-pnp-js'

import * as strings from 'EventPageWebPartStrings'
import { clientTag } from '../../shared/global';

export interface IExecutiveAssetsPageWebPartProps {
  description: string;
}

export default class ExecutiveAssetsPageWebPart extends BaseClientSideWebPart<IExecutiveAssetsPageWebPartProps> {
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
    const element: React.ReactElement<IExecutiveAssetsPageWebPartProps> = React.createElement(
      ExecutiveAssetsPage,
      {
        description: this.properties.description,
        context: this.context.pageContext
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion11(): Version {
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
