/* eslint-disable @typescript-eslint/no-var-requires */
import { SpinButton } from "@fluentui/react";
import { DatePicker, Dropdown, IDatePickerStyles, IDropdownOption, IDropdownStyles, ISpinButtonStyles, Label, Position, TextField } from "office-ui-fabric-react";
import * as React from "react";
import { IFieldEditorProps } from "./IFieldEditorProps";
import { IModernTaxonomyPickerProps, ModernTaxonomyPicker } from "@pnp/spfx-controls-react";
import { ITermData } from "./ITermData";

export default class FieldEditor extends React.Component<IFieldEditorProps, {}> {

  private getStringValue(value: string | undefined): string {

    if (value)
      return value.toString();
    else
      return "";
  
  }
  
  private getYesNoValue(value: string | undefined): string {

    switch (value) {
      
      case "Yes":
        return "Yes";

      default:
        return "No";
    }
  
  }

  private getDateValue(value: string | undefined): Date | undefined {

    if (value)
      return new Date(value);
    else
      return undefined;

  }

  private getDateStringValue(value: Date | null | undefined): string {

    if (value)
      return value.toString();
    else
      return "";
  
  }
  
  private onTrigger(fieldName: string, newValue?: string): void {
  
    this.props.callback(fieldName, this.getStringValue(newValue), true);
  
  }
   
  private onDropdownTrigger(fieldName: string, option?: IDropdownOption): void {
  
    this.props.callback(fieldName, this.getStringValue(option?.key.toString()), option?.selected);
  
  }
   
  private onDateTrigger(fieldName: string, newValue: Date | null | undefined): void {
  
    this.props.callback(fieldName, this.getDateStringValue(newValue), true);

  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private onTaxPickerChange(fieldName: string, terms: any): void {
    
    this.props.callback(fieldName, this.getStringValue(JSON.stringify(terms)), true);

  }

  public render(): React.ReactElement<IFieldEditorProps> {
    const {
      field,
      value,
      context
    } = this.props;

    switch (this.props.field.TypeDisplayName) {

      case "Single line of text": {
        return <TextField label={field.Title} defaultValue={value} onChange={(event, newValue?) => { this.onTrigger(field.InternalName, newValue) }} />
      }
      case "Choice": {

        const options: IDropdownOption[] = [];
        
        // Get the field choices
        if (field.Choices) {
          // Loop through the field choices
          field.Choices.forEach((choice: string) => {
            // Add it to the dropdown options
            options.push({key: choice, text: choice});
          });
        }

        const isMultiChoice:boolean = (field.TypeAsString === "MultiChoice");

        if (!isMultiChoice)
          return <Dropdown label={field.Title} options={options} defaultSelectedKey={this.getStringValue(value)} onChange={(event, option?, index?) => { this.onDropdownTrigger(field.InternalName, option) }} />
        else {
          let defaultChoices:string[] = [];
          if (value && value !== "") {
            defaultChoices = JSON.parse(value);
          }
          return <Dropdown multiSelect label={field.Title} options={options} defaultSelectedKeys={defaultChoices} onChange={(event, option?, index?) => { this.onDropdownTrigger(field.InternalName, option) }} />
        }
      }
      case "Number": {
        const spinButtonStyles: Partial<ISpinButtonStyles> = { spinButtonWrapper: { width: 300 } };
        return <SpinButton label={field.Title} labelPosition={Position.top} styles={spinButtonStyles} defaultValue={this.getStringValue(value)} onChange={(event, newValue?) => { this.onTrigger(field.InternalName, newValue) }} />
      }
      case "Yes/No": {
        const options: IDropdownOption[] = [
          { key: 'Yes', text: 'Yes' },
          { key: 'No', text: 'No' },
        ];
        const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 75 } };

        return <Dropdown label={field.Title} options={options} styles={dropdownStyles} defaultSelectedKey={this.getYesNoValue(value)} onChange={(event, option?, index?) => { this.onDropdownTrigger(field.InternalName, option) }} />
      }
      case "Date and Time": {
        const datePickerStyles: Partial<IDatePickerStyles> = { root: { maxWidth: 300 }};
        return <DatePicker label={field.Title} allowTextInput ariaLabel="Select a date" styles={datePickerStyles} value={this.getDateValue(value)} onSelectDate={(date: Date | null | undefined) => { this.onDateTrigger(field.InternalName, date )}} />
      }
      case "Managed Metadata": {

        // Load the XPath and DOM parsers
        const xpath = require('xpath');
        const dom = require('@xmldom/xmldom').DOMParser;
        
        // Read in the field schema as Xml
        const xml = new dom().parseFromString(field.SchemaXml, 'text/xml');

        // Get the Term Set Id value
        const TermSetId: string = xpath.select("//Property[Name/text()='TermSetId']/Value", xml)[0].firstChild.data;

        // Create a partial Term with the term information
        type PartialTermInfo = IModernTaxonomyPickerProps["initialValues"];
        let initialTerms: PartialTermInfo = [];

        // If we have a value
        if (value !== "") {

          // If this is a multi-value taxonomy field
          if (field.TypeAsString === "TaxonomyFieldTypeMulti") {

            // Get the term information
            const TermData = JSON.parse(value) as ITermData[];

            // Loop through the terms
            TermData.forEach(term => {

              // Create the initial term value
              const td = {
                id: term.TermID,
                labels: [
                  {
                    isDefault: true,
                    languageTag: 'en-US',
                    name: term.Label
                  }
                ]
              }

              // Add the term to our array
              initialTerms?.push(td);

            })
          }
          else {

            // Get the term information
            const TermData = JSON.parse(value) as ITermData;

            // Create the initial term
            initialTerms = [
              {
                id: TermData.TermID,
                labels: [
                  {
                    isDefault: true,
                    languageTag: 'en-US',
                    name: TermData.Label
                  }
                ]
              }
            ];
          }
        }

        return <ModernTaxonomyPicker allowMultipleSelections={(field.TypeAsString === "TaxonomyFieldTypeMulti")}
                  termSetId={TermSetId}
                  panelTitle="Select Term"
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context={context as any}
                  label={field.Title}
                  initialValues={initialTerms}
                  onChange={(terms) => { this.onTaxPickerChange(field.InternalName, terms) }}
               />
      }
      default: {
        return <div><Label>{field.Title}</Label><span>The &apos;{field.TypeDisplayName}&apos; field type is not yet supported.</span></div>;
      }
    }
  }
}
