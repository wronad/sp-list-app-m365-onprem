import { Spinner } from "@fluentui/react";
import { IPagedDataProvider } from "mgwdev-m365-helpers/lib/dal/dataProviders";
import * as React from "react";
import { COURSE_CODE, COURSE_NAME, IListItem, ISpListItem, SP_LIST_URL, extractSpListItems } from "../model/IListItem";
import { ListItemsGrid } from "./ListItemsGrid";
import { IHttpClient } from "mgwdev-m365-helpers/lib/dal/http";

export interface IListItemsProps {
    dataProvider?: IPagedDataProvider<ISpListItem>;
    graphClient?: IHttpClient;
}

export function ListItems(props: IListItemsProps) {
    const [items, setItems] = React.useState<IListItem[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);
    const [count, setCount] = React.useState<number>(0);

    React.useEffect(() => {
        props.dataProvider.getData().then((data) => {
            const listItems = extractSpListItems(data);
            setItems(listItems);
            setLoading(false);
        });
    }, []);

    React.useEffect(() => {
        if (count > 0) {
            props.dataProvider.getData().then((data) => {
                const listItems = extractSpListItems(data);
                setItems(listItems);
                setLoading(false);
            });
        }
    }, [count]);

    function addItem() {
        if (props.graphClient) {
            console.log("adding item to list...");
            const plus1 = count + 1;
            const postItem = {
                fields: {
                    Title: `${COURSE_NAME} - ${plus1}`,
                    COURSE_CODE: `${COURSE_CODE} - ${plus1}`,
                    COURSE_FREQUENCY: 'Card Holder',
                    TARGET_AUDIENCE: 'Initial'
                }
            }

            // const itemsArr: any[] = [ postItem ];

            const spOpts = {
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(postItem)
            }
            // console.log("posting: ", spOpts.body)
            // props.aadClient.post(SP_LIST_URL, AadHttpClient.configurations.v1, spOpts)
            props.graphClient.post(SP_LIST_URL, spOpts)
                .then(resp => {
                    console.log(JSON.stringify(resp));
                    setCount(plus1);
                })
                .catch(err => 
                {
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
        <button type="button" onClick={() => {addItem()}}>Add Item to SP List</button>
    </>);
}