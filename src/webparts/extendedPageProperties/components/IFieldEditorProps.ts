import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFieldInfo } from "@pnp/sp/fields";

export interface IFieldEditorProps {
    field: IFieldInfo;
    value: string;
    context: WebPartContext;
    callback: (fieldName: string, value: string, replaceValue: boolean | undefined) => void;
  }