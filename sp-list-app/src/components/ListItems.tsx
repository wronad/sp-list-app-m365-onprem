import { Spinner } from "@fluentui/react";
import { IPagedDataProvider } from "mgwdev-m365-helpers/lib/dal/dataProviders";
import { IHttpClient } from "mgwdev-m365-helpers/lib/dal/http";
import * as React from "react";
import {
  IListItem,
  IListItemPayloadOnline,
  extractSpListItems,
  mockNewListItem,
} from "../model/IListItem";
import {
  addListItemOnPrem,
  addListItemOnline,
  getListItemsOnPremAxios,
  getListItemsOnline,
} from "../services/SpListService";
import { ListItemsGrid } from "./ListItemsGrid";

export interface IListItemsProps {
  spOnlineDataProvider?: IPagedDataProvider<IListItemPayloadOnline>;
  spOnlineClient?: IHttpClient;
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
    if (props.spOnlineClient) {
      fetchDataOnline();
    } else {
      fetchDataOnPrem();
    }
  };

  const fetchDataOnline = () => {
    getListItemsOnline(props.spOnlineDataProvider).then((spListItems) => {
      itemsFetched(spListItems, true);
    });
  };

  const fetchDataOnPrem = () => {
    getListItemsOnPremAxios().then((spListItems) => {
      itemsFetched(spListItems, false);
    });
  };

  const itemsFetched = (spListItems: any, spOnline: boolean) => {
    if (spListItems) {
      const listItems = extractSpListItems(spListItems, spOnline);
      setItems(listItems);
    }
  };

  const addItem = () => {
    const plus1 = count + 1;
    const mockItem = mockNewListItem(plus1);
    if (props.spOnlineClient) {
      addItemOnline(mockItem, plus1);
    } else {
      addItemOnPrem(mockItem, plus1);
    }
  };

  const addItemOnline = (item: IListItem, num: number) => {
    addListItemOnline(props.spOnlineClient, item)
      .then((resp) => handleItemAdded(num))
      .catch((err) => console.log(JSON.stringify(err)));
  };

  const addItemOnPrem = (item: IListItem, num: number) => {
    addListItemOnPrem(item)
      .then((resp) => handleItemAdded(num))
      .catch((err) => console.log(JSON.stringify(err)));
  };

  const handleItemAdded = (num: number) => {
    setLoading(true);
    setCount(num);
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
