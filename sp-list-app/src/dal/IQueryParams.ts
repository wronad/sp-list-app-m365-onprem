export interface IQueryParams {
  listName: string;
  id?: number;
  listType?: string;
  columnName?: string;
  columnType?: string;
  itemData?: any;
  viewName?: string;
  rawQuery?: string;
  select?: string[];
  expand?: string;
  filter?: string; // where clause
  fileName?: string;
  file?: string | Blob | ArrayBuffer;
}
