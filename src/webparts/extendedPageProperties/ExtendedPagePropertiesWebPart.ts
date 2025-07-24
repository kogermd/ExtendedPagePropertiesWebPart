import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ExtendedPagePropertiesWebPartStrings';
import ExtendedPageProperties from './components/ExtendedPageProperties';
import { IExtendedPagePropertiesProps } from './components/IExtendedPagePropertiesProps';
import { propertyItem } from '../PropertyItem';
import { getGraph, getSP } from '../../pnpjs-config';
import { SPFI } from '@pnp/sp';
import { IFieldInfo } from '@pnp/sp/fields';
import "@pnp/sp/fields/list";
import { IRenderListDataParameters } from '@pnp/sp/lists';
import { IModernTaxonomyPickerProps } from '@pnp/spfx-controls-react';

export interface IExtendedPagePropertiesWebPartProps {
  title: string;
}

export default class ExtendedPagePropertiesWebPart extends BaseClientSideWebPart<IExtendedPagePropertiesWebPartProps> {

  private readonly docLibTitle = strings.SitePagesTitle;
  private _sp: SPFI;
  
    public async render(): Promise<void> {
  
      if (this.displayMode === DisplayMode.Edit) {
        const items: propertyItem[] = await this.getCurrentPageProperties();

        const element: React.ReactElement<IExtendedPagePropertiesProps> =
          React.createElement(ExtendedPageProperties, {
            title: this.properties.title,
            items,
            context: this.context,
            callback: this.setCurrentPageProperties.bind(this)
          });
  
        ReactDom.render(element, this.domElement);
      }
    }
  
    protected async onInit(): Promise<void> {
      await super.onInit();
      
      // Initialize our _sp object that we can then use in other packages without having to pass around the context.
      this._sp = getSP(this.context);
      getGraph(this.context);

    }
  
    getPropertyValue(property: propertyItem): string | number | boolean | Date {
  
      switch (property.field?.TypeDisplayName) {
  
        case "Single line of text": {
          return `"` + property.value + `"`;
        }
        case "Choice": {

          // If the field is a multi-choice field
          if (property.field?.TypeAsString === "MultiChoice") {
            return property.value;
          }
          else {
            // Single choice field, return the value;
            return `"` + property.value + `"`;
          }
        }
        case "Number": {
          return property.value;
        }
        case "Yes/No": {
          switch (property.value) {
      
            case "Yes":
              return "true";
      
            default:
              return "false";
          }
        }
        case "Date and Time": {
          return `"` + (new Date(property.value)).toISOString() + `"`;
        }
        case "Managed Metadata": {
          
          // Create a partial Term with the term information
          type PartialTermInfo = IModernTaxonomyPickerProps["initialValues"];
          const terms: PartialTermInfo = JSON.parse(property.value);
          let termValue: string = "";

          // Loop through the terms
          terms?.forEach(term => {

            // And add them to our return value
            termValue = termValue + term.labels[0].name + `|` + term.id;

            // If the control is a multi-value taxonomy field, add a semicolon separator
            if (property.field?.TypeAsString === "TaxonomyFieldTypeMulti") {
              termValue = termValue + `;`
            }

          });

          return `"` + termValue + `"`;
          
        }
        default: {
          return `"` + property.value + `"`;
        }
      }
      
    }
  
    setCurrentPageProperties(pageProperties: propertyItem[]): void {
  
      const currentPageId = this.context.pageContext.listItem?.id || -1;
      const page = this._sp.web.lists.getByTitle(this.docLibTitle).items.getById(currentPageId);

      let updateProps:string = `{`;
  
      pageProperties.forEach(property => {
  
        if (property.value && property.field?.TypeDisplayName !== "Multiple lines of text") {
  
          // Get the separator to use
          const separator = (updateProps === `{`) ? `"` : `, "`;
  
          // Create the update properties as a JSON string
          updateProps = updateProps + separator + property.field?.InternalName + `": ` + this.getPropertyValue(property);
 
        }
  
      });
  
      // Create the JSON object
      updateProps = updateProps + `}`;

      const hash = JSON.parse(updateProps); 
  
      page.update(hash)
          .then(() => {
            alert("Settings saved successfully!");
          })
          .catch((response) => {
            console.warn(response);
          });
      
    }
  
    protected async getCurrentPageProperties(): Promise<propertyItem[]> {
  
      // Get available fields in the Document Library
      const availableFields = await this._sp.web.lists
        .getByTitle(this.docLibTitle)
        .fields();
  
      // Filter to only non hidden fields
      const filteredAvailableFields: Map<string, IFieldInfo> = new Map();
      for (let i = 0; i < availableFields.length; i++) {
        const field = availableFields[i];
  
        // Don't get hidden, readonly, computed, file fields or fields that are not set to show in the Edit Form
        if (field.Hidden || field.ReadOnlyField || field.TypeDisplayName === "Computed" || field.TypeDisplayName === "File" || field.SchemaXml.indexOf("ShowInEditForm=\"FALSE\"") > -1) continue;
  
        // Add the field to the list
        filteredAvailableFields.set(field.InternalName, field);
      }
  
      const propertyItems: propertyItem[] = [];
      const currentPageId = this.context.pageContext.listItem?.id || -1;

      const renderListDataParams: IRenderListDataParameters = {
        ViewXml: "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Number'>" + currentPageId + "</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>",
      };

      const currentPageProperties = await this._sp.web.lists
        .getByTitle(this.docLibTitle)
        .renderListDataAsStream(renderListDataParams);

        filteredAvailableFields.forEach(field => {

          let value = "";

          if (currentPageProperties.LastRow > 0) {

            value = currentPageProperties.Row[0][field.InternalName];

            if (field.TypeDisplayName === "Managed Metadata") { // && field.TypeAsString === "TaxonomyFieldTypeMulti") {

              const hiddenField = availableFields.filter(f => f.Title === (field.Title + "_0"))[0];
              field.InternalName = hiddenField.InternalName;

            }

          }

          propertyItems.push({
            field,
            label: field.Title,
            value: getStringValue(value),
          });
    
        });
      
      return propertyItems;

    }
  
    protected onDispose(): void {
      ReactDom.unmountComponentAtNode(this.domElement);
    }
  
    protected get dataVersion(): Version {
      return Version.parse("1.0");
    }
  
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
        pages: [
          {
            groups: [
              {
                groupFields: [
                  PropertyPaneTextField("title", {
                    label: "Title",
                  }),
                ],
              },
            ],
          },
        ],
      };
    }
  }

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function getStringValue(value: any): string {
  
  switch (typeof(value)) {
    case "string":
      return value;

    case "object":
      return JSON.stringify(value);

    default:
      return value;

    }

}
  