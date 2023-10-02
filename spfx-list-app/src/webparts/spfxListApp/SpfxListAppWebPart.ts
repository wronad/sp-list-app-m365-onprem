import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "SpfxListAppWebPartStrings";
import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IListItemsProps,
  ListItems,
} from "sp-list-app/lib/components/ListItems";
import {
  ISpListItemPayload,
  MS_GRAPH_SP_LIST_FIELDS,
} from "sp-list-app/lib/model/IListItem";
import {
  IHttpClient,
  SPFxGraphHttpClient,
} from "sp-list-app/node_modules/mgwdev-m365-helpers";
import {
  GraphODataPagedDataProvider,
  IPagedDataProvider,
} from "sp-list-app/node_modules/mgwdev-m365-helpers/lib/dal";

export interface ISpfxListAppWebPartProps {
  description: string;
}

export default class SpfxListAppWebPart extends BaseClientSideWebPart<ISpfxListAppWebPartProps> {
  private _dataProvider: IPagedDataProvider<ISpListItemPayload>;
  private _graphClient: IHttpClient;

  protected async onInit(): Promise<void> {
    const aadHttpClient = await this.context.aadHttpClientFactory.getClient(
      "https://graph.microsoft.com"
    );
    const spfxGraphHttpClient = new SPFxGraphHttpClient(aadHttpClient);
    this._dataProvider = new GraphODataPagedDataProvider(
      spfxGraphHttpClient,
      MS_GRAPH_SP_LIST_FIELDS,
      true
    );
    this._graphClient = spfxGraphHttpClient;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public render(): void {
    const element: React.ReactElement<IListItemsProps> = React.createElement(
      ListItems,
      {
        spOnlineDataProvider: this._dataProvider,
        spOnlineClient: this._graphClient,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
