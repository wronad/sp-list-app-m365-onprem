import { SPFI } from "@pnp/sp";
import { IQueryParams } from "./IQueryParams";
import "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/presets/all";

// queries all the rows of a sharepoint list
export const getListItems = (
  spfi: SPFI,
  params: IQueryParams
): Promise<any> => {
  return spfi.web.lists
    .getByTitle(params.listName)
    .items.filter(params.filter)
    .expand(params.expand)
    .select(...params.select) // todo
    .top(10000)
    .getPaged();
};

/// creates a new sharepoint list item
export const addListItem = (spfi: SPFI, params: IQueryParams): Promise<any> => {
  return spfi.web.lists.getByTitle(params.listName).items.add(params.itemData);
};
