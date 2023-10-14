import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IListItemsProps,
  ListItems,
} from "sp-list-app/lib/components/ListItems";
import { SpfxSpHttpClient } from "sp-list-app/lib/spOnlineRestApi";

export interface ISpfxListAppWebPartProps {
  description: string;
}

export default class SpfxListAppWebPart extends BaseClientSideWebPart<ISpfxListAppWebPartProps> {
  private _restApiClient: SpfxSpHttpClient;

  protected async onInit(): Promise<void> {
    this._restApiClient = new SpfxSpHttpClient(this.context.spHttpClient);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public render(): void {
    const element: React.ReactElement<IListItemsProps> = React.createElement(
      ListItems,
      { spOnlineRestApi: this._restApiClient }
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
