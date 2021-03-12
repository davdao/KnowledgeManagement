import { Log, ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IComboBoxOption } from 'office-ui-fabric-react';
import { Constants } from '../../Constants';
import ISharePointManagedProperty from '../../model/search/ISharePointManagedProperty';
import { ISharePointSearchResponse } from '../../model/search/ISharePointSearchResponse';

export interface ISharePointDataSourceServices {
    getAvailableProperties(): Promise<IComboBoxOption[]>;
}

export class SharePointDataSourceServices implements ISharePointDataSourceServices {
        //Create a ServiceKey which will be used to consume the service.
        public static readonly serviceKey: ServiceKey<ISharePointDataSourceServices> = ServiceKey.create<ISharePointDataSourceServices>('aerow:ISharePointDataSourceServices', SharePointDataSourceServices);

        private _pageContext: PageContext;
        private _sphttpClient: SPHttpClient;

        constructor(serviceScope: ServiceScope) {
            serviceScope.whenFinished(() => {
                this._pageContext = serviceScope.consume<PageContext>(PageContext.serviceKey);
                this._sphttpClient = serviceScope.consume(SPHttpClient.serviceKey);
            });
        }

        public async getAvailableProperties(): Promise<IComboBoxOption[]> {
            let availableManagedPropertiesToReturn: IComboBoxOption[] = [];

            let managedProperties: ISharePointManagedProperty[] = [];
            let searchEndpointUrl = `${this._pageContext.web.absoluteUrl}/_api/search/postquery`;

            const postBody = {
                request:{ 
                    '__metadata': {
                        'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
                    },
                    'Querytext': '*',
                    'Refiners': 'ManagedProperties(filter=50000/0/*,sort=name/ascending)',
                    'RowLimit': 1
                }
            };

            try {
       
                const response = await this._sphttpClient.post(searchEndpointUrl, SPHttpClient.configurations.v1, {
                    body: JSON.stringify(postBody),
                    headers: {
                        'odata-version': '3.0',
                        'accept': 'application/json;odata=nometadata',
                        'X-ClientService-ClientTag': Constants.X_CLIENTSERVICE_CLIENTTAG,
                        'UserAgent': Constants.X_CLIENTSERVICE_CLIENTTAG
                    }
                });

                if (response.ok) {
                    const searchResponse: ISharePointSearchResponse = await response.json();
                    const refinementResultsRows = searchResponse.PrimaryQueryResult.RefinementResults;
                    const refinementRows: any = refinementResultsRows ? refinementResultsRows.Refiners : [];

                    // Map refinement results
                    refinementRows.forEach((refiner) => {
                        refiner.Entries.forEach((item) => {
                            managedProperties.push({
                                name: item.RefinementName
                            });
                        });
                    });
                }
            }
            catch (error) {
                Log.error("[SharePointSearchService.getAvailableManagedProperties()]", error);
                throw error;
            }

            availableManagedPropertiesToReturn = managedProperties.map(managedProperty => {
                return {
                    key: managedProperty.name,
                    text: managedProperty.name,
                } as IComboBoxOption;
            });

            return availableManagedPropertiesToReturn;
        }
/*
        private async getAllMetadataSearchProperties(): Promise<any> {

            let managedProperties: ISharePointManagedProperty[] = [];
            let searchEndpointUrl = `${this._pageContext.web.absoluteUrl}/_api/search/postquery`;

            const postBody = {
                request:{ 
                    '__metadata': {
                        'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
                    },
                    'Querytext': '*',
                    'Refiners': 'ManagedProperties(filter=50000/0/*,sort=name/ascending)',
                    'RowLimit': 1
                }
            };

            try {
       
                const response = await this._sphttpClient.post(searchEndpointUrl, SPHttpClient.configurations.v1, {
                    body: JSON.stringify(postBody),
                    headers: {
                        'odata-version': '3.0',
                        'accept': 'application/json;odata=nometadata',
                        'X-ClientService-ClientTag': Constants.X_CLIENTSERVICE_CLIENTTAG,
                        'UserAgent': Constants.X_CLIENTSERVICE_CLIENTTAG
                    }
                });

                if (response.ok) {
                    const searchResponse: ISharePointSearchResponse = await response.json();
                    const refinementResultsRows = searchResponse.PrimaryQueryResult.RefinementResults;
                    const refinementRows: any = refinementResultsRows ? refinementResultsRows.Refiners : [];

                    // Map refinement results
                    refinementRows.forEach((refiner) => {
                        refiner.Entries.forEach((item) => {
                            managedProperties.push({
                                name: item.RefinementName
                            });
                        });
                    });
                }
            }
            catch (error) {
                Log.error("[SharePointSearchService.getAvailableManagedProperties()]", error);
                throw error;
            }
            return managedProperties;
        }
*/
    }