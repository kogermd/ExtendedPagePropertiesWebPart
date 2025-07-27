/* eslint-disable @typescript-eslint/no-explicit-any */
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
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http-base';

export interface IExtendedPagePropertiesWebPartProps {
  title: string;
}

export default class ExtendedPagePropertiesWebPart extends BaseClientSideWebPart<IExtendedPagePropertiesWebPartProps> {

  private readonly docLibTitle = strings.SitePagesTitle;
  private _sp: SPFI;
  
  private sharedLockId: string | undefined = undefined;

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
          return property.value.replace(/","/gi, ";#").replace("[", "").replace("]", "")
        }
        else {
          // Single choice field, return the value;
          return `"` + property.value + `"`;
        }
      }
      case "Number": {
        return `"` + property.value + `"`;
      }
      case "Yes/No": {
        switch (property.value) {
    
          case "Yes":
            return `"true"`;
    
          default:
            return `"false"`;
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

    // Get the endpoint for the update method of the current page
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${this.context.pageContext.list?.id.toString()}')/items(${this.context.pageContext.listItem?.id})/validateupdatelistitem`;     

    let updateProps:string = `[{`;

    pageProperties.forEach(property => {

      if (property.value && property.field?.TypeDisplayName !== "Multiple lines of text") {

        // Get the separator to use
        const separator = (updateProps === `[{`) ? `` : `}, {`;

        // Create the update properties as a JSON string
        updateProps = updateProps + separator + `"FieldName": "` + property.field?.InternalName + `", "FieldValue": ` + this.getPropertyValue(property);

      }

    });
   
    // Close ther properties
    updateProps = updateProps + `}]`;

    // If we have properties
    if (updateProps !== `[{}]`) {

      // Convert the properties to a JSON object
      const updatedValues = JSON.parse(updateProps); 

      // Get the shared lock ID used for co-authoring
      this.getSharedLockId().then(() => {

        // Update the page with the shared lock ID
        updatePage(this.context.spHttpClient, endpoint, updatedValues, this.sharedLockId)
          .then(() => {
            alert("Settings saved successfully!");
          })
          .catch((response) => {
            console.warn(response);
          });
      }).catch((response) => {
        console.warn(response);
      });

    }
    
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

          if (field.TypeDisplayName === "Managed Metadata") {

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
  
  protected async getSharedLockId(): Promise<void> {

    // If we don't have the shared lock ID
    if (this.sharedLockId === undefined) {
    
      // Get the site pages component
      const sitePagesComponent = await SPComponentLoader.loadComponentById<any>("b6917cb1-93a0-4b97-a84d-7cf49975d4ec");

      // Get the shared lock ID from the site pages component
      if (sitePagesComponent?.PageStore?.fields) {
        this.sharedLockId = sitePagesComponent.PageStore.fields.SharedLockId;
      }

    }

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

async function updatePage(spHttpClient: SPHttpClient, endpointUrl: string, updatedValues: any[], sharedLockId: string | undefined): Promise<string> {

  try {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-Type', 'application/json');
    requestHeaders.append('Accept', 'application/json;odata=verbose');
    requestHeaders.append('OData-Version', '');

    // Create the body of the POST
    const body = {
      bNewDocumentUpdate: false,
      checkInComment: null,
      sharedLockId: sharedLockId,
      formValues: updatedValues,
      datesInUTC: true
    }

    const optionsWithData: ISPHttpClientOptions = {
      headers: requestHeaders,
      body: JSON.stringify(body)
    }

    // Submit the POST to update the page properties and await the response
    const response: SPHttpClientResponse = await spHttpClient.post(endpointUrl, SPHttpClient.configurations.v1, optionsWithData);

    // Get the response as a JSON object
    const responseJson = await response.json();
    const json = responseJson?.d ?? responseJson;
    const result = json?.ValidateUpdateListItem?.results[0] ?? json.error;

    // Return the result
    return JSON.stringify(result);

  } catch (error) {

    // If we got an error, return the error
    throw new Error(error);

  }
}