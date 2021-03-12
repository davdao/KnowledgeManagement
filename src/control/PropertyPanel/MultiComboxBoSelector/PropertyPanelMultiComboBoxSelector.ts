import { ServiceScope } from '@microsoft/sp-core-library';
import { IPropertyPaneCustomFieldProps, IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import { IComboBoxOption } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import MultiComboBoxSelector from './MultiComboBoxSelector';

export interface IPropertyPanelMultiComboBoxSelectorInternalProps extends IPropertyPanelMultiComboBoxSelectorProps, IPropertyPaneCustomFieldProps {

}

export interface IPropertyPanelMultiComboBoxSelectorProps {
    label: string;
    availableOptions?: Promise<IComboBoxOption[]>;
    selectedKey?: string | number;
    disabled?: boolean;
}

export class PropertyPanelMultiComboBoxSelector implements IPropertyPaneField<IPropertyPanelMultiComboBoxSelectorProps> {

    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public shouldFocus?: boolean;
    public properties: IPropertyPanelMultiComboBoxSelectorInternalProps;
    private elem: HTMLElement;

    constructor(targetProperty: string, properties: IPropertyPanelMultiComboBoxSelectorProps) {
        this.targetProperty = targetProperty;
        
        this.properties = {
            key: properties.label,
            availableOptions: properties.availableOptions,
            label: properties.label,
            selectedKey: properties.selectedKey ? properties.selectedKey : null,
            disabled: properties.disabled ? properties.disabled : null,
            onRender: this.onRender.bind(this),
            onDispose: this.onDispose.bind(this)
        };
    }    
    
    private onDispose(element: HTMLElement): void {
        ReactDom.unmountComponentAtNode(element);
    }

    private onRender(elem: HTMLElement): void {
        if (!this.elem) {
            this.elem = elem;
        }

        const element: React.ReactElement<any> = React.createElement(MultiComboBoxSelector, {
            label: this.properties.label,
            disabled: this.properties.disabled,
            availableOptions: this.properties.availableOptions
        });

        ReactDom.render(element, elem);
    }
}