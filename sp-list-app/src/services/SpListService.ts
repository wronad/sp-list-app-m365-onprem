import axios from "axios";
import { IPagedDataProvider } from "mgwdev-m365-helpers/lib/dal/dataProviders";
import {
  IHttpClient,
  IHttpClientResponse,
} from "mgwdev-m365-helpers/lib/dal/http";
import {
  IListItem,
  ISpListItemPayload,
  MS_GRAPH_SP_LIST,
  bundleBodyForOnline,
  bundleDataForOnPrem,
} from "../model/IListItem";

// SP MS 365 / Online /////////////////////////////////////////////////////////////////

const SP_OPTS = {
  headers: {
    "Content-Type": "application/json",
  },
  body: "",
};

export const getListItemsOnline = async (
  dataProvider: IPagedDataProvider<ISpListItemPayload>
): Promise<ISpListItemPayload[]> => {
  return dataProvider.getData();
};

export const addListItemOnline = async (
  spOnlineClient: IHttpClient,
  listItem: IListItem
): Promise<IHttpClientResponse> => {
  SP_OPTS.body = bundleBodyForOnline(listItem);
  return spOnlineClient.post(MS_GRAPH_SP_LIST, SP_OPTS);
};

// SP On Prem / Subscription Edition (SE) ////////////////////////////////////////////

const LIST_NAME = "ListAppExample";

const SITE_URL = "https://8r1bcm.sharepoint.com";
const LISTS_URL = SITE_URL + "/_api/web/lists";
const SP_LIST = `${LISTS_URL}/GetByTitle('${LIST_NAME}')/items`;
const ITEM_TYPE = `SP.Data.${SP_LIST}ListItem`;

const axiosCfg = {
  headers: {
    accept: "application/json;odata=verbose",
    "content-type": "application/json;odata=verbose",
  },
};

const spListApi = axios.create({
  baseURL: SITE_URL,
});

export const getListItemsOnPrem = async () => {
  return spListApi
    .get(SP_LIST, axiosCfg)
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

export const addListItemOnPrem = async (listItem: IListItem) => {
  const data = bundleDataForOnPrem(listItem, ITEM_TYPE);
  return spListApi
    .post(SP_LIST, data, axiosCfg)
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
