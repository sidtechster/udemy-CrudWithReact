import { IListItem } from "./IListItem";

export interface ICrudWithReactState {
    status: string;
    ListItem: IListItem;
    ListItems: IListItem[];
}