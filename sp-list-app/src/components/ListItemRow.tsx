import * as React from "react";
import { IListItem } from "../model/IListItem";

export interface IListItemRowProps {
  row: IListItem;
}

export function ListItemRow(props: IListItemRowProps) {
  return (
    <tr key={props.row.id}>
      <td>{props.row.id}</td>
      <td>{props.row.courseName}</td>
      <td>{props.row.courseCode}</td>
      <td>{props.row.courseFrequency}</td>
      <td>{props.row.targetAudience}</td>
    </tr>
  );
}
