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
  addListItemOnLine,
  getListItemsOnLine,
  getListItemsOnPrem,
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

  const fetchDataOnline = async () => {
    getListItemsOnLine(props.spOnlineDataProvider).then((spListItems) => {
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
        const stop = 1;
      }
    });
  };

  const fetchData = async () => {
    if (props.spOnlineClient) {
      fetchDataOnline();
    } else {
      fetchDataOnPrem();
    }
  };

  React.useEffect(() => {
    fetchData();
  }, []);

  React.useEffect(() => {
    if (count > 0) {
      fetchData();
    }
  }, [count]);

  const addItem = () => {
    if (props.spOnlineClient) {
      console.log("adding item to list...");
      const plus1 = count + 1;
      const mockItem = mockNewListItem(plus1);
      addListItemOnLine(props.spOnlineClient, mockItem)
        .then((resp) => {
          setCount(plus1);
        })
        .catch((err) => {
          console.log(JSON.stringify(err));
        });
    }
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
