import axios from "axios";
import { cfg } from "../app-config";
import {
  IListItem,
  bundleDataForOnPrem,
  bundleDataForOnlineApi,
} from "../model/IListItem";
import {
  IHttpClientResponse as IHttpClientResponseApi,
  SpfxSpHttpClient,
} from "../spOnlineRestApi";

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
const SITE_URL = `${urlPrefix}${cfg.SP_SITE}`;
const SP_LIST = `${SITE_URL}/_api/web/lists/GetByTitle('${LIST_NAME}')/items`;

// SP MS 365 / Online /////////////////////////////////////////////////////////////////

const SP_OPTS = {
  headers: {
    "X-RequestDigest": "",
  },
  body: "",
};

export const getListItemsApiOnline = async (
  spOnlineApi: SpfxSpHttpClient
): Promise<IHttpClientResponseApi> => {
  return spOnlineApi.get(SP_LIST).then((response) => {
    return response.json();
  });
};

export const addListItemApiOnline = async (
  spOnlineApi: SpfxSpHttpClient,
  listItem: IListItem
): Promise<IHttpClientResponseApi> => {
  return getContextDigest().then((digest) => {
    SP_OPTS.headers["X-RequestDigest"] = digest;
    SP_OPTS.body = bundleDataForOnlineApi(listItem);
    return spOnlineApi.post(SP_LIST, SP_OPTS);
  });
};

// SP On Prem / Subscription Edition (SE) ////////////////////////////////////////////

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
