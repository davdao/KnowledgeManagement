import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'KnowledgeManagementWebPartWebPartStrings';
import KnowledgeManagementWebPart from './components/KnowledgeManagementWebPart';

export interface IKnowledgeManagementWebPartWebPartProps {

}

export default class KnowledgeManagementWebPartWebPart extends BaseClientSideWebPart<IKnowledgeManagementWebPartWebPartProps> {

  public render(): void {

    const element: React.ReactElement = React.createElement(
      KnowledgeManagementWebPart,
      {
        spfxContext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
