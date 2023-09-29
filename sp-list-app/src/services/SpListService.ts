import axios from "axios";
import { IPagedDataProvider } from "mgwdev-m365-helpers/lib/dal/dataProviders";
import {
  IHttpClient,
  IHttpClientResponse,
} from "mgwdev-m365-helpers/lib/dal/http";
import {
  IListItem,
  ISpListItemPayload,
  SP_LIST_URL,
  bundleBody,
} from "../model/IListItem";

export const SITE_URL = "https://8r1bcm.sharepoint.com";
export const LISTS_URL = SITE_URL + "/_api/web/lists";
const LIST_NAME = "ListAppExample";

const axiosCfg = {
  headers: {
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
  },
};

const spListApi = axios.create({
  baseURL: SITE_URL,
});

const SP_OPTS = {
  headers: {
    "Content-Type": "application/json",
  },
  body: "",
};

// SP MS 365 / Online /////////////////////////////////////////////////////////////////

export const getListItemsOnLine = async (
  dataProvider: IPagedDataProvider<ISpListItemPayload>
): Promise<ISpListItemPayload[]> => {
  return dataProvider.getData();
};

export const addListItemOnLine = async (
  spOnlineClient: IHttpClient,
  listItem: IListItem
): Promise<IHttpClientResponse> => {
  SP_OPTS.body = bundleBody(listItem);
  return spOnlineClient.post(SP_LIST_URL, SP_OPTS);
};

// SP On Prem / Subscription Edition (SE) ////////////////////////////////////////////

export const getListItemsOnPrem = async () => {
  return spListApi
    .get(`${LISTS_URL}/GetByTitle('${LIST_NAME}')`, axiosCfg)
    .then((resp) => {
      console.log("Response");
      console.log(resp);
      return resp;
    })
    .catch((err) => {
      console.log(err);
      return null;
    });
};

// export const AddListItemsOnPrem = async (
// ): Promise<ISpListItem[]> => {
//   // return dataProvider.getData();
// };
