import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SpfxListAppWebPartStrings';
import { ISpListItem, MS_GRAPH_URL_SP_SITE } from "sp-list-app/lib/model/IListItem";
import { IListItemsProps, ListItems } from "sp-list-app/lib/components/ListItems";
import { GraphODataPagedDataProvider, IPagedDataProvider } from 'sp-list-app/node_modules/mgwdev-m365-helpers/lib/dal';
import {IHttpClient, SPFxGraphHttpClient } from 'sp-list-app/node_modules/mgwdev-m365-helpers';

export interface ISpfxListAppWebPartProps {
  description: string;
}

export default class SpfxListAppWebPart extends BaseClientSideWebPart<ISpfxListAppWebPartProps> {
  private _dataProvider: IPagedDataProvider<ISpListItem>;
  private _graphClient: IHttpClient;

  protected async onInit(): Promise<void> {
    const aadHttpClient = await this.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    const spfxGraphHttpClient = new SPFxGraphHttpClient(aadHttpClient);
    this._dataProvider = new GraphODataPagedDataProvider(spfxGraphHttpClient, MS_GRAPH_URL_SP_SITE, true);
    this._graphClient = spfxGraphHttpClient;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public render(): void {
    const element: React.ReactElement<IListItemsProps> = React.createElement(
      ListItems,
      {
        dataProvider: this._dataProvider,
        graphClient: this._graphClient,
      }
    );

    ReactDom.render(element, this.domElement);
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