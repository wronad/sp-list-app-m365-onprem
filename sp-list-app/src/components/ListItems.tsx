import { Spinner } from "@fluentui/react";
import { IPagedDataProvider } from "mgwdev-m365-helpers/lib/dal/dataProviders";
import { IHttpClient } from "mgwdev-m365-helpers/lib/dal/http";
import * as React from "react";
import {
  IListItem,
  ISpListItemPayload,
  extractSpListItems,
  mockNewListItem,
} from "../model/IListItem";
import {
  addListItemOnPrem,
  addListItemOnline,
  getListItemsOnPrem,
  getListItemsOnline,
} from "../services/SpListService";
import { ListItemsGrid } from "./ListItemsGrid";

export interface IListItemsProps {
  spOnlineDataProvider?: IPagedDataProvider<ISpListItemPayload>;
  spOnlineClient?: IHttpClient;
}

export const ListItems = (props: IListItemsProps) => {
  const [items, setItems] = React.useState<IListItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [count, setCount] = React.useState<number>(0);

  React.useEffect(() => {
    fetchData();
  }, []);

  React.useEffect(() => {
    if (count > 0) {
      fetchData();
    }
  }, [count]);

  const fetchData = async () => {
    if (props.spOnlineClient) {
      fetchDataOnline();
    } else {
      fetchDataOnPrem();
    }
  };

  const fetchDataOnline = async () => {
    getListItemsOnline(props.spOnlineDataProvider).then((spListItems) => {
      if (spListItems) {
        const listItems = extractSpListItems(spListItems);
        setItems(listItems);
        setLoading(false);
      }
    });
  };

  const fetchDataOnPrem = async () => {
    getListItemsOnPrem().then((spListItems) => {
      if (spListItems) {
        const debugThis = 1;
        console.log(JSON.stringify(spListItems));
      }
    });
  };

  const addItem = () => {
    console.log("adding item to list...");
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
      .then((resp) => {
        setCount(num);
      })
      .catch((err) => {
        console.log(JSON.stringify(err));
      });
  };

  const addItemOnPrem = (item: IListItem, num: number) => {
    addListItemOnPrem(item)
      .then((resp) => {
        const debugThis = 2;
        console.log(JSON.stringify(resp));
        setCount(num);
      })
      .catch((err) => {
        console.log(JSON.stringify(err));
      });
  };

  if (loading) {
    return <Spinner />;
  }
  return (
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
  );
};
