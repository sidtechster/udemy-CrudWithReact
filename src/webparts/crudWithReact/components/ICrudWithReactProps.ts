import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICrudWithReactProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
