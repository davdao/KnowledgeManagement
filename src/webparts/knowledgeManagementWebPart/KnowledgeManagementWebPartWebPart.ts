import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'KnowledgeManagementWebPartWebPartStrings';
import KnowledgeManagementWebPart from './components/KnowledgeManagementWebPart';
import { IKnowledgeManagementWebPartProps } from './components/IKnowledgeManagementWebPartProps';
import { PropertyPanelMultiComboBoxSelector } from '../../control/PropertyPanel/MultiComboxBoSelector/PropertyPanelMultiComboBoxSelector';
import { ISharePointDataSourceServices, SharePointDataSourceServices } from '../../service/SharePointServices/SharePointDataSourceServices';
import { IComboBoxOption } from 'office-ui-fabric-react/lib/components/ComboBox/ComboBox.types';

export interface IKnowledgeManagementWebPartWebPartProps {
  description: string;
  _availableManagedProperties: IComboBoxOption[];
}

export default class KnowledgeManagementWebPartWebPart extends BaseClientSideWebPart<IKnowledgeManagementWebPartWebPartProps> {

  private _sharePointDataSourceServices: ISharePointDataSourceServices;

  public render(): void {
    const element: React.ReactElement<IKnowledgeManagementWebPartProps> = React.createElement(
      KnowledgeManagementWebPart,
      {
        description: this.properties.description,
        serviceScope: this.context.serviceScope
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    this._sharePointDataSourceServices = this.context.serviceScope.consume(SharePointDataSourceServices.serviceKey);    
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                new PropertyPanelMultiComboBoxSelector('selectedProperties', {
                  label: "selectedProperties",
                  availableOptions: this.getAllProperties(),
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async getAllProperties(){
    if(this._sharePointDataSourceServices)
      return await this._sharePointDataSourceServices.getAvailableProperties()
    
  }
}
