import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'KnowledgeManagementWebPartWebPartStrings';
import KnowledgeManagementWebPart from './components/KnowledgeManagementWebPart';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneChoiceGroup, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { sp } from "@pnp/sp/presets/all";
import { ThemeBackgrounds } from '../../common/Enum';
import { IDataSource } from '../../models/dataSource/IDataSource';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { SharePointSearchDataSource } from '../../dataSource/SharePointSearchDataSource';
import { ServiceScopeHelper } from '../../helpers/ServiceScopeHelper';
import { BaseDataSource } from '../../models/dataSource/BaseDataSource';
import { TokenService } from '../../services/tokenService/TokenService';
import { PropertyPanelSearchMetadataProperty } from '../../controls/PropertyPanel/SearchMetadataSelector/PropertyPanelSearchMetadataProperty';
import { IComboBoxOption } from 'office-ui-fabric-react';

export interface IKnowledgeManagementWebPartWebPartProps {
  title: any;
  ThemeLayout: ThemeBackgrounds;
  ShowHideSearchBar: boolean;
  ShowHideRefinementPanel: boolean;
  availableManagedProperties: IComboBoxOption[];
}

export default class KnowledgeManagementWebPartWebPart extends BaseClientSideWebPart<IKnowledgeManagementWebPartWebPartProps> {  

    /**
   * The service scope for this specific Web Part instance
   */
  private webPartInstanceServiceScope: ServiceScope;
  
    /**
   * The selected data source for the WebPart
   */
  private dataSource: IDataSource;

  protected async onInit(): Promise<void> {
    // Initializes the Web Part instance services
    this.initializeWebPartServices();
  }

  public async render(): Promise<void> {
    this.dataSource = await this.GetDataSource(); 
      /*       
    let tes = await this.dataSource.getAvailableProperties();
console.log(tes);*/
    const element: React.ReactElement = React.createElement(
      KnowledgeManagementWebPart,
      {
        context: this.context,
        Theme: this.properties.ThemeLayout ? this.properties.ThemeLayout : ThemeBackgrounds.List,
        ShowHideSearchBar: this.properties.ShowHideSearchBar ? this.properties.ShowHideSearchBar : false,
        ShowHideRefinementPanel: this.properties.ShowHideRefinementPanel ? this.properties.ShowHideRefinementPanel : false,
        webPartTitleProps: {
          displayMode: this.displayMode,
          title: this.properties.title,
          updateProperty: (value: string) => {
            this.properties.title = value;
          }
        }
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
          groups: [
            {
              groupName: strings.PropertyPanel.DisplayDescription,
              groupFields: [
                PropertyPaneToggle('ShowHideSearchBar', {
                  label: strings.PropertyPanel.HideShowSearchBar,
                  onText: strings.PropertyPanel.ToggleResultYes,
                  offText: strings.PropertyPanel.ToggleResultNo,
                }),
                PropertyPaneToggle('ShowHideRefinementPanel', {
                  label: strings.PropertyPanel.HideShowRefinementPanel,
                  onText: strings.PropertyPanel.ToggleResultYes,
                  offText: strings.PropertyPanel.ToggleResultNo,
                }),
                PropertyPaneChoiceGroup('ThemeLayout', {
                  label: strings.PropertyPanel.ThemeLayout.ThemeDescription,
                  options: [
                    {
                      key: ThemeBackgrounds.List,
                      text: strings.PropertyPanel.ThemeLayout.ListView,
                      iconProps: { officeFabricIconFontName: "AllApps" },
                      checked: true
                    },
                    {
                      key: ThemeBackgrounds.Brick,
                      text: strings.PropertyPanel.ThemeLayout.GridView,
                      iconProps: { officeFabricIconFontName: "waffle" }
                    }
                  ]
                }),
                new PropertyPanelSearchMetadataProperty('MetadataSearchProperties', {
                  label:"bonjour",
                  onLoadOptions: this.GetAllSearchProperties.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private  GetAllSearchProperties() {
    return this.dataSource.getAvailableProperties()
  }
  private initializeWebPartServices(): void {

    // Register specific Web Part service instances
    this.webPartInstanceServiceScope = this.context.serviceScope.startNewChild();
    
    this.webPartInstanceServiceScope.finish();
  }

  private async GetDataSource(): Promise<IDataSource> {
    let dataSource: IDataSource = undefined;
    let serviceKey: ServiceKey<IDataSource> = undefined;

    serviceKey = ServiceKey.create<IDataSource>('ModernSearch:SharePointSearchDataSource', SharePointSearchDataSource);
    
    return new Promise<IDataSource>((resolve, reject) => {

          // Register here services we want to expose to custom data sources (ex: TokenService)
          // The instances are shared across all data sources. It means when properties will be set once for all consumers. Be careful manipulating these instance properties. 
          const childServiceScope = ServiceScopeHelper.registerChildServices(this.webPartInstanceServiceScope, [
            serviceKey          
          ]);

          childServiceScope.whenFinished(async () => {
            // Register the data source service in the Web Part scope only (child scope of the current scope)
            dataSource = childServiceScope.consume<IDataSource>(serviceKey);

            // Verifiy if the data source implements correctly the IDataSource interface and BaseDataSource methods
            const isValidDataSource = (dataSourceInstance: IDataSource): dataSourceInstance is BaseDataSource<any> => {
              return (
                (dataSourceInstance as BaseDataSource<any>).getAppliedFilters !== undefined &&
                (dataSourceInstance as BaseDataSource<any>).getData !== undefined &&
                (dataSourceInstance as BaseDataSource<any>).getFilterBehavior !== undefined &&
                (dataSourceInstance as BaseDataSource<any>).getItemCount !== undefined &&
                (dataSourceInstance as BaseDataSource<any>).getPropertyPaneGroupsConfiguration !== undefined &&
                (dataSourceInstance as BaseDataSource<any>).getTemplateSlots !== undefined &&
                (dataSourceInstance as BaseDataSource<any>).onInit !== undefined &&
                (dataSourceInstance as BaseDataSource<any>).onPropertyUpdate !== undefined
              );
            };

            if (!isValidDataSource(dataSource)) {
              reject(new Error(strings.InvalidDataSourceInstance));
            }

            // Initialize the data source with current Web Part properties
            if (dataSource) {
                  // Initializes Web part lifecycle methods and properties
               //   dataSource.properties = this.properties.dataSourceProperties;
                  dataSource.context = this.context;
                  dataSource.render = this.render;
            }

            // Initializes available services
            dataSource.serviceKeys = {
              TokenService: TokenService.ServiceKey
            };

            await dataSource.onInit();

          /*  // Initialize slots
            if (isEmpty(this.properties.templateSlots)) {
              this.properties.templateSlots = dataSource.getTemplateSlots();
              this._defaultTemplateSlots = dataSource.getTemplateSlots();
            }*/
            resolve(dataSource);
          });      
    });
  }
}
