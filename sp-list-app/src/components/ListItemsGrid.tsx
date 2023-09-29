import * as React from "react";
import { IListItem } from "../model/IListItem";
import { ListItemRow } from "./ListItemRow";

export interface IListItemGridProps {
  rows: IListItem[];
}

export const ListItemsGrid = (props: IListItemGridProps) => {
  return (
    <table>
      <tr>
        <td>ID</td>
        <td>Course Name</td>
        <td>Course Code</td>
        <td>Course Frequency</td>
        <td>Target Audience</td>
      </tr>
      <>
        {props.rows.map((row) => (
          <ListItemRow row={row} />
        ))}
      </>
    </table>
  );
};
