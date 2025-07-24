import * as React from 'react';
import styles from './ExtendedPageProperties.module.scss';
import type { IExtendedPagePropertiesProps } from './IExtendedPagePropertiesProps';
import { propertyItem } from '../../PropertyItem';
import { PrimaryButton } from 'office-ui-fabric-react';
import FieldEditor from './FieldEditor';

export default class ExtendedPageProperties extends React.Component<IExtendedPagePropertiesProps, {}> {

  private updateSettings(): void {
    
    this.props.callback(this.props.items);

  }

  private handleCallback(fieldName: string, value: string, addValue: boolean | undefined): void {

    // Get the page property for the field that was passed
    const pageProperties:propertyItem[] = this.props.items.filter(x => x.field && x.field.InternalName === fieldName);

    // If we found the page property, update the value
    if (pageProperties.length > 0) {

      // Get the first (only) page property
      const pageProperty:propertyItem = pageProperties[0];

      // Check to see if this is a multi-choice field
      const isMultiChoice:boolean = (pageProperty.field?.TypeAsString === "MultiChoice");
      if (isMultiChoice) {

        // Get the choices for the field
        let choices:string[] = [];
        
        if (pageProperty.value)
          choices = JSON.parse(pageProperty.value);

        // If we are adding the value, add it to our array
        choices = addValue ? [...choices, value] : choices.filter(choice => choice !== value);

        // Update the array as a string
        pageProperty.value = JSON.stringify(choices);
      }
      else {
        // Not a multi-choice field, set the value
        pageProperty.value = value;
      }

    }
    
  }
  
  public render(): React.ReactElement<IExtendedPagePropertiesProps> {

    const {
      title,
      items,
      context
    } = this.props;

    return (
      <section className={styles.extendedPageProperties}>
        <h2>{title}</h2>
        <div className={styles.content}>
          {items.map((item: propertyItem) => (
            <div key={item.field?.Id} className={styles.item}>
              {item.field ? (
                <FieldEditor field={item.field} value={item.value} callback={this.handleCallback.bind(this)} context={context} />
              ) : (
                <span />
              )}
            </div>
          ))}
        </div>
        <br/><br/>
        <div>
          <PrimaryButton text="Update Settings" onClick={this.updateSettings.bind(this)} />
        </div>
      </section>
    )
  }
}
