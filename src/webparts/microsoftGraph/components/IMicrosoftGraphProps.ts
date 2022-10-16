import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ClientMode } from "./ClientMode";

export interface IMicrosoftGraphProps {
  clientMode: ClientMode;
  context: WebPartContext;
}
