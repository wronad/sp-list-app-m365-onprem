import axios from "axios";
import { IPagedDataProvider } from "mgwdev-m365-helpers/lib/dal/dataProviders";
import {
  IHttpClient,
  IHttpClientResponse,
} from "mgwdev-m365-helpers/lib/dal/http";
import { cfg } from "../app-config";
import {
  IListItem,
  IListItemPayloadOnline,
  bundleBodyForOnline,
  bundleDataForOnPrem,
} from "../model/IListItem";

// example sp sites
// SP_SITE = "8r1bcm.sharepoint.com"; // online dev tenant
// SP_SITE = "sp-onprem:3110/dev-site"; // onPrem MS Azure
// SP_SITE = "soceur.sof.socom.smil.mil/ppws/sandbox/ReactApps"; // onPrem

// example howto get SP list id
// SP -> navigate to list -> settings -> list settigs -> RSS settings
// redirects to:
//   https://${SP_SITE}/_layouts/15/listedit.aspx?List=%7B53eff35b-3e5a-4ef3-b87f-3baad80b982a%7D
// LIST_ID = "53eff35b-3e5a-4ef3-b87f-3baad80b982a";

// SP MS 365 / Online /////////////////////////////////////////////////////////////////

export const MS_GRAPH = "https://graph.microsoft.com";
export const MS_GRAPH_SP_LIST = `${MS_GRAPH}/v1.0/sites/${cfg.SP_SITE}/lists/${cfg.LIST_ID}/items`;

const SP_OPTS = {
  headers: {
    "Content-Type": "application/json",
  },
  body: "",
};

export const getListItemsOnline = async (
  dataProvider: IPagedDataProvider<IListItemPayloadOnline>
): Promise<IListItemPayloadOnline[]> => {
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
let urlPrefix = "https://";
if (!cfg.SSL) {
  urlPrefix = "http://"; // Azure dev env
}
const SITE_URL = `${urlPrefix}${cfg.SP_SITE}`;
const SP_LIST = `${SITE_URL}/_api/web/lists/GetByTitle('${LIST_NAME}')/items`;

// http://${SP_SITE}/_api/web/lists/getbytitle('ListAppExample')/ListItemEntityTypeFullName
const ITEM_TYPE = `SP.Data.${LIST_NAME}ListItem`;

const axiosCfg = {
  headers: {
    accept: "application/json;odata=verbose",
    "content-type": "application/json;odata=verbose",
  },
};

const spListApi = axios.create({
  baseURL: SITE_URL,
});

export const getListItemsOnPremAxios = async () => {
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

const getContextDigest = async () => {
  return spListApi
    .post(`${SITE_URL}/_api/contextinfo`, axiosCfg)
    .then((resp) => {
      return resp.data.FormDigestValue;
    })
    .catch((err) => {
      console.error(err);
      return "";
    });
};

export const addListItemOnPrem = async (listItem: IListItem) => {
  return getContextDigest().then((digest) => {
    axiosCfg.headers["X-RequestDigest"] = digest;
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
  });
};
