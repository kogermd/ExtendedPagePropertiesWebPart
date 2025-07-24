import { WebPartContext } from "@microsoft/sp-webpart-base";
import { propertyItem } from "../../PropertyItem";

export interface IExtendedPagePropertiesProps {
  title: string;
  items: propertyItem[];
  context: WebPartContext;
  callback: (pageProperties: propertyItem[]) => void
}
