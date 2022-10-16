import { IUserItem } from "./IUserItem";

export interface IMicrosoftGraphState {
  users: Array<IUserItem>;
  searchFor: string;
}