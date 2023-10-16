import axios from "axios";
import { cfg } from "../app-config";
import {
  IListItem,
  bundleDataForOnPrem,
  bundleDataForOnlineApi,
  bundleItem,
} from "../model/IListItem";
import { SpfxSpHttpClient } from "../dal";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import { addListItem, getListItems } from "../dal/spPnpDal";
import { IQueryParams } from "../dal/IQueryParams";

// example sp sites
// SP_SITE = "######.sharepoint.com"; // online dev tenant
// SP_SITE = "sp-onprem:3110/dev-site"; // onPrem MS Azure
// SP_SITE = "soceur.*/ppws/sandbox/ReactApps"; // onPrem

// example howto get SP list id for ListAppExample list
// SP -> navigate to list -> settings -> list settigs -> RSS settings
// redirects to:
//   https://${SP_SITE}/_layouts/15/listedit.aspx?List=%7BCB76740F-10ED-40B2-BE51-523C2F02E9DE%7D
//   LIST_ID = "CB76740F-10ED-40B2-BE51-523C2F02E9DE";

const LIST_NAME = "ListAppExample";
let urlPrefix = "https://";
if (!cfg.SSL) {
  urlPrefix = "http://"; // Azure dev env
}
export const SITE_URL = `${urlPrefix}${cfg.SP_SITE}`;
const SP_LIST_ENDPOINT = `${SITE_URL}/_api/web/lists/GetByTitle('${LIST_NAME}')/items`;

const getContextDigest = async () => {
  return AXIOS_SP_LIST_API.post(`${SITE_URL}/_api/contextinfo`, AXIOS_CFG)
    .then((resp) => {
      return resp.data.FormDigestValue;
    })
    .catch((err) => {
      console.error("SPListService.getContextDigest", err);
      return "";
    });
};

// Only for SPFx, can NOT use for standalone react app ////////////////////////////////////////////

const SP_OPTS = {
  headers: {
    "X-RequestDigest": "",
  },
  body: "",
};
const VAL = "value";

// SPFx client - REST API
export const getListItemsSpfxClient = async (
  spfxRestClient: SpfxSpHttpClient
): Promise<any> => {
  return spfxRestClient
    .get(SP_LIST_ENDPOINT)
    .then((response) => response.json())
    .then((resp) => {
      if (VAL in resp) {
        return resp[VAL];
      }
      return [];
    })
    .catch((err) => {
      console.log("SpListService.getListItemsSpfxClient", err);
      return [];
    });
};

// SPFx client - REST API
export const addListItemSpfxClient = async (
  spfxRestClient: SpfxSpHttpClient,
  listItem: IListItem
): Promise<number> => {
  return getContextDigest().then((digest) => {
    SP_OPTS.headers["X-RequestDigest"] = digest;
    SP_OPTS.body = bundleDataForOnlineApi(listItem);
    return spfxRestClient
      .post(SP_LIST_ENDPOINT, SP_OPTS)
      .then((response) => {
        return 1;
      })
      .catch((err) => {
        console.log("SpListService.addListItemSpfxClient", err);
        return 0;
      });
  });
};

// SP PnP client
export const getListItemsSpPnp = async (spPnpClient: SPFI): Promise<any> => {
  const params: IQueryParams = {
    listName: LIST_NAME,
    select: [],
    expand: "",
    filter: "",
  };
  return getListItems(spPnpClient, params)
    .then((response) => {
      if (response?.results) {
        return response.results;
      }
      return [];
    })
    .catch((err) => {
      console.log("SpListService.getListItemsSpPnp", err);
      return [];
    });
};

// SP PnP client
export const addListItemSpPnp = async (
  spPnpClient: SPFI,
  listItem: IListItem
): Promise<number> => {
  const item = bundleItem(listItem);
  const params: IQueryParams = {
    listName: LIST_NAME,
    itemData: item,
  };
  return addListItem(spPnpClient, params)
    .then((response) => (response?.data?.Id ? response.data.Id : 0))
    .catch((err) => {
      console.log("SpListService.addListItemSpPnp", err);
      return 0;
    });
};

// Ok for SP Online, OnPrem / Subscription Edition (SE), & standalone apps ///////////////

// http://${SP_SITE}/_api/web/lists/getbytitle('ListAppExample')/ListItemEntityTypeFullName
const ITEM_TYPE = `SP.Data.${LIST_NAME}ListItem`;

const AXIOS_CFG = {
  headers: {
    accept: "application/json;odata=verbose",
    "content-type": "application/json;odata=verbose",
  },
};

const AXIOS_SP_LIST_API = axios.create({
  baseURL: SITE_URL,
});

// axios SP List API
export const getListItemsRestApi = async () => {
  return AXIOS_SP_LIST_API.get(SP_LIST_ENDPOINT, AXIOS_CFG)
    .then((resp) => {
      if (resp?.data?.d?.results?.length) {
        return resp.data.d.results;
      }
    })
    .catch((err) => {
      console.error("SPListService.getListItemsOnPrem", err);
      return [];
    });
};

// axios SP List API
export const addListItemRestApi = async (
  listItem: IListItem
): Promise<number> => {
  return getContextDigest().then((digest) => {
    AXIOS_CFG.headers["X-RequestDigest"] = digest;
    const data = bundleDataForOnPrem(listItem, ITEM_TYPE);
    return AXIOS_SP_LIST_API.post(SP_LIST_ENDPOINT, data, AXIOS_CFG)
      .then((response) => (response?.data?.d?.Id ? response.data.d.Id : 0))
      .catch((err) => {
        console.error("SPListService.getContextDigest", err);
        return 0;
      });
  });
};
