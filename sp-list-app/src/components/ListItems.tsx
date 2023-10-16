import { Spinner } from "@fluentui/react";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import * as React from "react";
import { cfg } from "../app-config";
import { SpfxSpHttpClient } from "../dal";
import {
  IListItem,
  extractSpListItems,
  mockNewListItem,
} from "../model/IListItem";
import {
  addListItemRestApi,
  addListItemSpPnp,
  addListItemSpfxClient,
  getListItemsRestApi,
  getListItemsSpPnp,
  getListItemsSpfxClient,
} from "../services/SpListService";
import { ListItemsGrid } from "./ListItemsGrid";

// gulpfile.js set-sp-site --api options:
const PnP = "pnp"; // default
const REST = "rest";
const SPFx = "spfx";

export interface IListItemsProps {
  spfxRestClient?: SpfxSpHttpClient;
  spPnpClient?: SPFI;
}

export const ListItems = (props: IListItemsProps) => {
  const [count, setCount] = React.useState<number>(0);
  const [items, setItems] = React.useState<IListItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    if (loading) {
      fetchData();
    }
  }, [loading]);

  React.useEffect(() => {
    setLoading(false);
  }, [items]);

  const fetchData = () => {
    switch (cfg.API_TYPE) {
      case PnP:
      default:
        fetchDataPnp();
        break;
      case REST:
        fetchDataRest();
        break;
      case SPFx:
        fetchDataSpfx();
        break;
    }
  };

  const fetchDataSpfx = () => {
    getListItemsSpfxClient(props.spfxRestClient).then((spListItems) =>
      itemsFetched(spListItems)
    );
  };

  const fetchDataPnp = () => {
    getListItemsSpPnp(props.spPnpClient).then((spListItems) =>
      itemsFetched(spListItems)
    );
  };

  const fetchDataRest = () => {
    getListItemsRestApi().then((spListItems) => itemsFetched(spListItems));
  };

  const itemsFetched = (spListItems: any) => {
    if (spListItems) {
      const listItems = extractSpListItems(spListItems);
      setItems(listItems);
    }
  };

  const addItem = () => {
    const plus1 = count + 1;
    const mockItem = mockNewListItem(plus1);
    switch (cfg.API_TYPE) {
      case PnP:
      default:
        addItemOnlineSpPnp(mockItem, plus1);
        break;
      case REST:
        addItemRestApi(mockItem, plus1);
        break;
      case SPFx:
        addItemSpfx(mockItem, plus1);
        break;
    }
  };

  const addItemSpfx = (item: IListItem, num: number) => {
    addListItemSpfxClient(props.spfxRestClient, item)
      .then((id) => handleItemAdded(id ? num : 0))
      .catch((err) => console.log(JSON.stringify(err)));
  };

  const addItemOnlineSpPnp = (item: IListItem, num: number) => {
    addListItemSpPnp(props.spPnpClient, item).then((id) =>
      handleItemAdded(id ? num : 0)
    );
  };

  const addItemRestApi = (item: IListItem, num: number) => {
    addListItemRestApi(item).then((id) => handleItemAdded(id ? num : 0));
  };

  const handleItemAdded = (num: number) => {
    setLoading(true);
    if (num) {
      setCount(num);
    }
  };

  return (
    <>
      {loading ? (
        <Spinner />
      ) : (
        <>
          <ListItemsGrid rows={items} />
          <br></br>
          <br></br>
          <br></br>
          <button
            type="button"
            onClick={() => {
              addItem();
            }}
          >
            Add Item to SP List
          </button>
        </>
      )}
    </>
  );
};
