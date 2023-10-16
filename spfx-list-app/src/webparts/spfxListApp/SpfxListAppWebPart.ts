import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IListItemsProps,
  ListItems,
} from "sp-list-app/lib/components/ListItems";
import { SpfxSpHttpClient } from "sp-list-app/lib/dal";
import { SITE_URL } from "sp-list-app/lib/services/SpListService";
import { getSpFrameworkIF } from "../../dal/spPnpDal";

export interface ISpfxListAppWebPartProps {
  description: string;
}

export default class SpfxListAppWebPart extends BaseClientSideWebPart<ISpfxListAppWebPartProps> {
  private _restApiClient: SpfxSpHttpClient;
  private _pnpApiClient: SPFI;

  protected async onInit(): Promise<void> {
    this._restApiClient = new SpfxSpHttpClient(this.context.spHttpClient);
    this._pnpApiClient = getSpFrameworkIF(this.context, SITE_URL);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public render(): void {
    const element: React.ReactElement<IListItemsProps> = React.createElement(
      ListItems,
      { spfxRestClient: this._restApiClient, spPnpClient: this._pnpApiClient }
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
            description: "SPFx react example list app ",
          },
          groups: [
            {
              groupName: "SPFx react app group",
              groupFields: [
                PropertyPaneTextField("description", {
                  label: "SPFx react app label",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
0;
