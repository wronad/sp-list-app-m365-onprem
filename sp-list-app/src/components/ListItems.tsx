import { Spinner } from "@fluentui/react";
import * as React from "react";
import {
  IListItem,
  extractSpListItems,
  mockNewListItem,
} from "../model/IListItem";
import {
  addListItemApiOnline,
  addListItemOnPrem,
  getListItemsApiOnline,
  getListItemsOnPrem,
} from "../services/SpListService";
import { SpfxSpHttpClient } from "../spOnlineRestApi";
import { ListItemsGrid } from "./ListItemsGrid";

export interface IListItemsProps {
  spOnlineRestApi?: SpfxSpHttpClient;
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
    if (props.spOnlineRestApi) {
      fetchDataOnline();
    } else {
      fetchDataOnPrem();
    }
  };

  const fetchDataOnline = () => {
    getListItemsApiOnline(props.spOnlineRestApi).then((spListItems) => {
      itemsFetched(spListItems, true);
    });
  };

  const fetchDataOnPrem = () => {
    getListItemsOnPrem().then((spListItems) => {
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
    if (props.spOnlineRestApi) {
      addItemOnline(mockItem, plus1);
    } else {
      addItemOnPrem(mockItem, plus1);
    }
  };

  const addItemOnline = (item: IListItem, num: number) => {
    addListItemApiOnline(props.spOnlineRestApi, item)
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
