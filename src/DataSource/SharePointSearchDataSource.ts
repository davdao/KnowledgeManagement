import * as React from 'react';
import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import { ServiceScope } from '@microsoft/sp-core-library';
import { isEmpty } from "@microsoft/sp-lodash-subset";
import { ISharePointSearchService } from "../services/searchService/ISharePointSearchService";
import { SharePointSearchService } from "../services/searchService/SharePointSearchService";
import LocalizationHelper from "../helpers/LocalizationHelper";
import { PageContext } from "@microsoft/sp-page-context";
import { TokenService } from "../services/tokenService/TokenService";
import { IComboBoxOption } from "office-ui-fabric-react";
import { ISharePointSearchQuery, SortDirection, ISort } from "../models/search/ISharePointSearchQuery";
import { DateHelper } from '../helpers/DateHelper';
import { DataFilterHelper } from '../helpers/DataFilterHelper';
import { ISortFieldConfiguration, SortFieldDirection } from '../models/search/ISortFieldConfiguration';
import { EnumHelper } from '../helpers/EnumHelper';
import { IDataContext } from '../models/dataSource/IDataContext';
import { IDataFilterConfiguration } from '../models/search/IDataFilterConfiguration';
import { FilterComparisonOperator, IDataFilter } from '../models/search/IDataFilter';
import { BuiltinTemplateSlots, ITemplateSlot } from '../Common/ITemplateSlot';
import { FilterBehavior } from '../models/dataSource/FilterBehavior';
import { IDataSourceData } from '../models/dataSource/IDataSourceData';
import { BaseDataSource } from '../models/dataSource/BaseDataSource';

export enum BuiltinSourceIds {
    Documents = 'e7ec8cee-ded8-43c9-beb5-436b54b31e84',
    ItemsMatchingContentType = '5dc9f503-801e-4ced-8a2c-5d1237132419',
    ItemsMatchingTag = 'e1327b9c-2b8c-4b23-99c9-3730cb29c3f7',
    ItemsRelatedToCurrentUser = '48fec42e-4a92-48ce-8363-c2703a40e67d',
    ItemsWithSameKeywordAsThisItem = '5c069288-1d17-454a-8ac6-9c642a065f48',
    LocalPeopleResults = 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
    LocalReportsAndDataResults = '203fba36-2763-4060-9931-911ac8c0583b',
    LocalSharePointResults = '8413cd39-2156-4e00-b54d-11efd9abdb89',
    LocalVideoResults = '78b793ce-7956-4669-aa3b-451fc5defebf',
    Pages = '5e34578e-4d08-4edc-8bf3-002acf3cdbcc',
    Pictures = '38403c8c-3975-41a8-826e-717f2d41568a',
    Popular = '97c71db1-58ce-4891-8b64-585bc2326c12',
    RecentlyChangedItems = 'ba63bbae-fa9c-42c0-b027-9a878f16557c',
    RecommendedItems = 'ec675252-14fa-4fbe-84dd-8d098ed74181',
    Wiki = '9479bf85-e257-4318-b5a8-81a180f5faa1',
}

/**
 * SharePoint search data source property pane properties
 */
export interface ISharePointSearchDataSourceProperties {

    /**
     * The search query template
     */
    queryTemplate: string;

    /**
     * SharePoint result source GUID
     */
    resultSourceId: string;

    /**
     * Flag indicating if the query rules should enabled/disabled
     */
    enableQueryRules: boolean;

    /**
     * Flag indicating if the OneDrive for Business results shoud be included/excluded
     */
    includeOneDriveResults: boolean;

    /**
     * The KQL or FQL refinement filters to apply to the query
     */
    refinementFilters: string;

    /**
     * Flag indicating if the query should be localized
     */
    enableLocalization: boolean;

    /**
     * The search query language to use (locale ID)
     */
    searchQueryLanguage: number;

    /**
     * The search managed properties to retrieve
     */
    selectedProperties: string[];

    /**
     * The sort fields configuration
     */
    sortList: ISortFieldConfiguration[];

    /**
     * Flag indicating if the audience targeting should be enabled
     */
    enableAudienceTargeting: boolean;
}

export class SharePointSearchDataSource extends BaseDataSource<ISharePointSearchDataSourceProperties> {

    private _availableLanguages: IPropertyPaneDropdownOption[] = [];
    private _availableManagedProperties: IComboBoxOption[] = [];
    private _resultSourcesOptions: IComboBoxOption[] = [];
    private _sharePointSearchService: ISharePointSearchService;
    private _pageContext: PageContext;
   // private _tokenService: ITokenService;
    private _currentLocaleId: number;
    
    /**
     * The data source items count
     */
    private _itemsCount: number;

    /*
    * A date helper instance
    */
    private dateHelper: DateHelper;

    /**
    * The moment.js library reference
    */
    private moment: any;

    public constructor(serviceScope: ServiceScope) {
        super(serviceScope);

        serviceScope.whenFinished(() => {
            this._sharePointSearchService = serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
            this._pageContext = serviceScope.consume<PageContext>(PageContext.serviceKey);
      //      this._tokenService = serviceScope.consume<ITokenService>(TokenService.ServiceKey);
        });
    }

    public async onInit(): Promise<void> {

        this.initProperties();

        this.dateHelper = this.serviceScope.consume<DateHelper>(DateHelper.ServiceKey);
        this.moment = await this.dateHelper.moment();

        this._currentLocaleId = LocalizationHelper.getLocaleId(this._pageContext.cultureInfo.currentUICultureName);

        // Initialize the list of available languages
        if (this._availableLanguages.length == 0) {
            const languages = await this._sharePointSearchService.getAvailableQueryLanguages();

            this._availableLanguages = languages.map(language => {
                return {
                    key: language.Lcid,
                    text: `${language.DisplayName} (${language.Lcid})`
                };
            });
        }
    }

    public async getData(dataContext: IDataContext): Promise<IDataSourceData> {

        const searchQuery = await this.buildSharePointSearchQuery(dataContext);
        const results = await this._sharePointSearchService.search(searchQuery);

        let data: IDataSourceData = {
            items: results.relevantResults,
            filters: results.refinementResults,
            queryModification: results.queryModification,
            secondaryResults: results.secondaryResults,
            spellingSuggestion: results.spellingSuggestion,
            promotedResults: results.promotedResults
        };

        this._itemsCount = results.totalRows;

        return data;
    }
/*
    public onCustomPropertyUpdate(propertyPath: string, newValue: any): void {

        if (propertyPath.localeCompare('dataSourceProperties.selectedProperties') === 0) {
            this.properties.selectedProperties = (cloneDeep(newValue) as IComboBoxOption[]).map(v => { return v.key as string; });
            this.context.propertyPane.refresh();
            this.render();
        }

        if (propertyPath.localeCompare('dataSourceProperties.resultSourceId') === 0) {
            this.properties.resultSourceId = (newValue as IComboBoxOption).key as string;
            this.context.propertyPane.refresh();
            this.render();
        }
    }*/

    public getFilterBehavior(): FilterBehavior {
        return FilterBehavior.Dynamic;
    }

    public getItemCount(): number {
        return this._itemsCount;
    }

    public getTemplateSlots(): ITemplateSlot[] {
        return [
            {
                slotName: BuiltinTemplateSlots.Title,
                slotField: 'Title'
            },
            {
                slotName: BuiltinTemplateSlots.Path,
                slotField: 'DefaultEncodingURL'
            },
            {
                slotName: BuiltinTemplateSlots.Summary,
                slotField: 'HitHighlightedSummary'
            },
            {
                slotName: BuiltinTemplateSlots.FileType,
                slotField: 'FileType'
            },
            {
                slotName: BuiltinTemplateSlots.PreviewImageUrl,
                slotField: 'AutoPreviewImageUrl' // Field added automatically
            },
            {
                slotName: BuiltinTemplateSlots.PreviewUrl,
                slotField: 'AutoPreviewUrl' // Field added automatically
            },
            {
                slotName: BuiltinTemplateSlots.Author,
                slotField: 'CreatedBy'
            },
            {
                slotName: BuiltinTemplateSlots.Tags,
                slotField: 'owstaxidmetadataalltagsinfo'
            },
            {
                slotName: BuiltinTemplateSlots.Date,
                slotField: 'Created'
            },
            {
                slotName: BuiltinTemplateSlots.SiteId,
                slotField: 'NormSiteID'
            },
            {
                slotName: BuiltinTemplateSlots.ListId,
                slotField: 'NormListID'
            },
            {
                slotName: BuiltinTemplateSlots.ItemId,
                slotField: 'NormUniqueID'
            },
            {
                slotName: BuiltinTemplateSlots.IsFolder,
                slotField: 'ContentTypeId'
            },
            {
                slotName: BuiltinTemplateSlots.PersonQuery,
                slotField: 'UserName'
            },
            {
                slotName: BuiltinTemplateSlots.UserDisplayName,
                slotField: 'Title'
            },
            {
                slotName: BuiltinTemplateSlots.UserEmail,
                slotField: 'UserName'
            }
        ];
    }

    private initProperties(): void {
       /* this.properties.queryTemplate = this.properties.queryTemplate ? this.properties.queryTemplate : "{searchTerms}";
        this.properties.enableQueryRules = this.properties.enableQueryRules !== undefined ? this.properties.enableQueryRules : false;
        this.properties.enableLocalization = this.properties.enableLocalization !== undefined ? this.properties.enableLocalization : false;
        this.properties.includeOneDriveResults = this.properties.includeOneDriveResults !== undefined ? this.properties.includeOneDriveResults : false;
        this.properties.refinementFilters = this.properties.refinementFilters ? this.properties.refinementFilters : '';
        this.properties.selectedProperties = this.properties.selectedProperties !== undefined ? this.properties.selectedProperties :
            [
                'Title',
                'Path',
                'DefaultEncodingURL',
                'FileType',
                'HitHighlightedSummary',
                'AuthorOWSUSER',
                'owstaxidmetadataalltagsinfo',
                'Created',
                'UniqueID',
                'NormSiteID',
                'NormListID',
                'NormUniqueID',
                'ContentTypeId',
                'UserName',
                'JobTitle',
                'WorkPhone',
                'SPSiteUrl',
                'SiteTitle',
                'CreatedBy'
            ];*/
/*        this.properties.resultSourceId = this.properties.resultSourceId !== undefined ? this.properties.resultSourceId : BuiltinSourceIds.LocalSharePointResults;
        this.properties.sortList = this.properties.sortList !== undefined ? this.properties.sortList : [];*/
    }

    private getBuiltinSourceIdOptions(): IComboBoxOption[] {

        this._resultSourcesOptions = [
            {
                key: BuiltinSourceIds.Documents,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.Documents)
            },
            {
                key: BuiltinSourceIds.ItemsMatchingContentType,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.ItemsMatchingContentType)
            },
            {
                key: BuiltinSourceIds.ItemsMatchingTag,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.ItemsMatchingTag)
            },
            {
                key: BuiltinSourceIds.ItemsRelatedToCurrentUser,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.ItemsRelatedToCurrentUser)
            },
            {
                key: BuiltinSourceIds.ItemsWithSameKeywordAsThisItem,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.ItemsWithSameKeywordAsThisItem)
            },
            {
                key: BuiltinSourceIds.LocalPeopleResults,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.LocalPeopleResults)
            },
            {
                key: BuiltinSourceIds.LocalReportsAndDataResults,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.LocalReportsAndDataResults)
            },
            {
                key: BuiltinSourceIds.LocalSharePointResults,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.LocalSharePointResults)
            },
            {
                key: BuiltinSourceIds.LocalVideoResults,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.LocalVideoResults)
            },
            {
                key: BuiltinSourceIds.Pages,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.Pages)
            },
            {
                key: BuiltinSourceIds.Pictures,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.Pictures)
            },
            {
                key: BuiltinSourceIds.Popular,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.Popular)
            },
            {
                key: BuiltinSourceIds.RecentlyChangedItems,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.RecentlyChangedItems)
            },
            {
                key: BuiltinSourceIds.RecommendedItems,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.RecommendedItems)
            },
            {
                key: BuiltinSourceIds.Wiki,
                text: EnumHelper.getEnumKeyByEnumValue(BuiltinSourceIds, BuiltinSourceIds.Wiki)
            },
        ];

        return this._resultSourcesOptions;
    }

    public async getAvailableProperties(): Promise<IComboBoxOption[]> {

        const searchManagedProperties = await this._sharePointSearchService.getAvailableManagedProperties();

        this._availableManagedProperties = searchManagedProperties.map(managedProperty => {
            return {
                key: managedProperty.name,
                text: managedProperty.name,
            } as IComboBoxOption;
        });

        return this._availableManagedProperties;
    }

    private _convertToSortList(sortList: ISortFieldConfiguration[]): ISort[] {
        return sortList.map(e => {

            let direction;

            switch (e.sortDirection) {
                case SortFieldDirection.Ascending:
                    direction = SortDirection.Ascending;
                    break;

                case SortFieldDirection.Descending:
                    direction = SortDirection.Descending;
                    break;

                default:
                    direction = SortDirection.Ascending;
                    break;
            }

            return {
                Property: e.sortField,
                Direction: direction
            } as ISort;
        });
    }

    public async buildSharePointSearchQuery(dataContext: IDataContext): Promise<ISharePointSearchQuery> {

        // Build the search query according to options
        let searchQuery: ISharePointSearchQuery = {};

        searchQuery.ClientType = 'PnPModernSearch';
        searchQuery.Properties = [{
            Name: "EnableDynamicGroups",
            Value: {
                BoolVal: true,
                QueryPropertyValueTypeIndex: 3
            }
        }, {
            Name: "EnableMultiGeoSearch",
            Value: {
                BoolVal: true,
                QueryPropertyValueTypeIndex: 3
            }
        }, {
            Name: "ClientFunction",
            Value: {
                StrVal: "PnPSearchWebPart",
                QueryPropertyValueTypeIndex: 1
            }
        }, {
            // Sample query: foo:test
            // As "foo" is not an OOB schema property it will be treated as text "foo test" instead
            // of non-existing property query - yielding results instead of a blank page
            Name: "ImplicitPropertiesAsStrings",
            Value: {
                BoolVal: true,
                QueryPropertyValueTypeIndex: 3
            }
        }];
        if (this._pageContext.list) {
            searchQuery.Properties.push({
                Name: "ListId",
                Value: {
                    StrVal: this._pageContext.list.id.toString(),
                    QueryPropertyValueTypeIndex: 1
                }
            });
        }

        if (this._pageContext.listItem) {
            searchQuery.Properties.push({
                Name: "ListItemId",
                Value: {
                    StrVal: this._pageContext.listItem.id.toString(),
                    QueryPropertyValueTypeIndex: 1
                }
            });
        }

        searchQuery.Querytext = dataContext.inputQueryText;

     //   searchQuery.EnableQueryRules = this.properties.enableQueryRules;
     //   searchQuery.QueryTemplate = await this._tokenService.resolveTokens(this.properties.queryTemplate);

     /*   if (this.properties.resultSourceId) {
            searchQuery.SourceId = this.properties.resultSourceId;
        }*/

        // Enable phoenetic search for people result source
        if (searchQuery.SourceId && searchQuery.SourceId.toLocaleLowerCase() === BuiltinSourceIds.LocalPeopleResults) {
            searchQuery.EnableNicknames = true;
            searchQuery.EnablePhonetic = true;
        } else {
            searchQuery.EnableNicknames = false;
            searchQuery.EnablePhonetic = false;
        }

 //       searchQuery.Culture = this.properties.searchQueryLanguage !== undefined && this.properties.searchQueryLanguage !== null ? this.properties.searchQueryLanguage : this._currentLocaleId;

        // Determine time zone bias
        let timeZoneBias = {
            WebBias: this._pageContext.legacyPageContext.webTimeZoneData.Bias,
            WebDST: this._pageContext.legacyPageContext.webTimeZoneData.DaylightBias,
            UserBias: null,
            UserDST: null,
            Id: this._pageContext.legacyPageContext.webTimeZoneData.Id
        };

        if (this._pageContext.legacyPageContext.userTimeZoneData) {
            timeZoneBias.UserBias = this._pageContext.legacyPageContext.userTimeZoneData.Bias;
            timeZoneBias.UserDST = this._pageContext.legacyPageContext.userTimeZoneData.DaylightBias;
            timeZoneBias.Id = this._pageContext.legacyPageContext.webTimeZoneData.Id;
        }

        searchQuery['TimeZoneId'] = timeZoneBias.Id;

   //     let refinementFilters: string[] = !isEmpty(this.properties.refinementFilters) ? [this.properties.refinementFilters] : [];

        if (!isEmpty(dataContext.filters)) {

            // Set list of refiners to retrieve
            searchQuery.Refiners = dataContext.filters.filtersConfiguration.map(filterConfig => {

                // Special case with Date managed properties
                const regexExpr = "(RefinableDate\\d+)(?=,|$)|" +
                    "(RefinableDateInvariant00\\d+)(?=,|$)|" +
                    "(RefinableDateSingle\\d+)(?=,|$)|" +
                    "(LastModifiedTime)(?=,|$)|" +
                    "(LastModifiedTimeForRetention)(?=,|$)|" +
                    "(Created)(?=,|$)|" +
                    "(Date\\d+)(?=,|$)|" +
                    "(EndDate)(?=,|$)|" +
                    "(.+OWSDATE)(?=,|$)|" +
                    "(EventsRollUpEndDate)(?=,|$)|" +
                    "(EventsRollUpStartDate)(?=,|$)|" +
                    "(FirstPublishedDate)(?=,|$)|" +
                    "(ImageDateCreated)(?=,|$)|" +
                    "(LastAnalyticsUpdateTime)(?=,|$)|" +
                    "(ModifierDates)(?=,|$)|" +
                    "(ClassificationLastScan)(?=,|$)|" +
                    "(ComplianceTagWrittenTime)(?=,|$)|" +
                    "(ContentModifiedTime)(?=,|$)|" +
                    "(DocumentAnalyticsLastActivityTimestamp)(?=,|$)|" +
                    "(ExpirationTime)(?=,|$)|" +
                    "(LastSharedByTime)(?=,|$)|" +
                    "(StartDate)(?=,|$)|" +
                    "(TagEventDate)(?=,|$)|" +
                    "(processingtime)(?=,|$)|" +
                    "(ExtractedDate)(?=,|$)";

                const refinableDateRegex = new RegExp(regexExpr.replace(/\s+/gi, ''), 'gi');
                if (refinableDateRegex.test(filterConfig.filterName)) {

                    const pastYear = this.moment(new Date()).subtract(1, 'years').subtract('minutes', 1).toISOString();
                    const past3Months = this.moment(new Date()).subtract(3, 'months').subtract('minutes', 1).toISOString();
                    const pastMonth = this.moment(new Date()).subtract(1, 'months').subtract('minutes', 1).toISOString();
                    const pastWeek = this.moment(new Date()).subtract(1, 'week').subtract('minutes', 1).toISOString();
                    const past24hours = this.moment(new Date()).subtract(24, 'hours').subtract('minutes', 1).toISOString();
                    const today = new Date().toISOString();

                    return `${filterConfig.filterName}(discretize=manual/${pastYear}/${past3Months}/${pastMonth}/${pastWeek}/${past24hours}/${today})`;

                } else {
                    return filterConfig.filterName;
                }

            }).join(',');

            // Get refinement filters
            if (dataContext.filters.selectedFilters.length > 0) {

        /*        // Make sure, if we have multiple filters, at least two filters have values to avoid apply an operator ('or','and') on only one condition failing the query.
                if (dataContext.filters.selectedFilters.length > 1 && dataContext.filters.selectedFilters.filter(selectedFilter => selectedFilter.values.length > 0).length > 1) {
                    const refinementString = this.buildRefinementQueryString(dataContext.filters.selectedFilters, dataContext.filters.filtersConfiguration).join(',');
                    if (!isEmpty(refinementString)) {
                        refinementFilters = refinementFilters.concat([`${dataContext.filters.filterOperator}(${refinementString})`]);
                    }

                } else {
                    refinementFilters = refinementFilters.concat(this.buildRefinementQueryString(dataContext.filters.selectedFilters, dataContext.filters.filtersConfiguration));
                }*/
            }

        }

     //   searchQuery.RefinementFilters = refinementFilters;

        // Paging settings
        searchQuery.RowLimit = dataContext.itemsCountPerPage ? dataContext.itemsCountPerPage : 50;

        if (dataContext.pageNumber === 1) {
            searchQuery.StartRow = 0;
        } else {
            searchQuery.StartRow = (dataContext.pageNumber - 1) * searchQuery.RowLimit;
        }

        searchQuery.TrimDuplicates = false;
  /*      searchQuery.SortList = this._convertToSortList(this.properties.sortList);
        searchQuery.SelectProperties = this.properties.selectedProperties;*/

        // Toggle to include user's personal OneDrive results as a secondary result block
        // https://docs.microsoft.com/en-us/sharepoint/support/search/private-onedrive-results-not-included
     /*   if (this.properties.includeOneDriveResults) {
            searchQuery.Properties.push({
                Name: "ContentSetting",
                Value: {
                    IntVal: 3,
                    QueryPropertyValueTypeIndex: 2
                }
            });
        }*/
        return searchQuery;
    }

    /**
     * Build the refinement condition in FQL format
     * @param selectedFilters The selected filter array
     * @param filtersConfiguration The current filters configuration
     * @param encodeTokens If true, encodes the taxonomy refinement tokens in UTF-8 to work with GET requests. Javascript encodes natively in UTF-16 by default.
     */
    private buildRefinementQueryString(selectedFilters: IDataFilter[], filtersConfiguration: IDataFilterConfiguration[], encodeTokens?: boolean): string[] {

        let refinementQueryConditions: string[] = [];

        selectedFilters.forEach(filter => {

            let operator: any = filter.operator;

            // Get the configuration for this filter
            const filterConfiguration: IDataFilterConfiguration = DataFilterHelper.getConfigurationForFilter(filter, filtersConfiguration);

            // The configuration should always be here for a filter. Not a valid scenario otherwise.
            if (filterConfiguration) {

                // Mutli values
                if (filter.values.length > 1) {

                    let startDate = null;
                    let endDate = null;

                    // A refiner can have multiple values selected in a multi or mon multi selection scenario
                    // The correct operator is determined by the refiner display template according to its behavior
                    const conditions = filter.values.map(filterValue => {

                        let value = filterValue.value;

                        if (this.moment(value, this.moment.ISO_8601, true).isValid()) {

                            if (!startDate && (filterValue.operator === FilterComparisonOperator.Geq || filterValue.operator === FilterComparisonOperator.Gt)) {
                                startDate = value;
                            }

                            if (!endDate && (filterValue.operator === FilterComparisonOperator.Lt || filterValue.operator === FilterComparisonOperator.Leq)) {
                                endDate = value;
                            }
                        }

                        // TODO A Etudier : We know the taxonomy picker sends the selected taxonomy ID every time so we can safely use the value without processing
                    /*    if (filterConfiguration.selectedTemplate === BuiltinFilterTemplates.TaxonomyPicker) {
                            value = `GP0|#${filterValue.value},L0|#0${filterValue.value}`; // Refine a SharePoint taxonomy term (only items with that specific term are retrieved)
                        }*/

                        return /ǂǂ/.test(value) && encodeTokens ? encodeURIComponent(value) : value;
                    });

                    if (startDate && endDate) {
                        refinementQueryConditions.push(`${filter.filterName}:range(${startDate},${endDate})`);
                    } else {
                        refinementQueryConditions.push(`${filter.filterName}:${operator}(${conditions.join(',')})`);
                    }

                } else {

                    // Single value
                    if (filter.values.length === 1) {

                        const filterValue = filter.values[0];

                        // See https://sharepoint.stackexchange.com/questions/258081/how-to-hex-encode-refiners/258161
                        let refinementToken = /ǂǂ/.test(filterValue.value) && encodeTokens ? encodeURIComponent(filterValue.value) : filterValue.value;

                        //TODO A Etudier :  We know the taxonomy picker sends the selected taxonomy ID every time so we can safely use the value without processing
                     /*   if (filterConfiguration.selectedTemplate === BuiltinFilterTemplates.TaxonomyPicker) {
                            refinementToken = `or(GP0|#${filterValue.value}, L0|#0${filterValue.value})`; // Refine a SharePoint taxonomy term (only results with that term). See https://docs.microsoft.com/en-us/sharepoint/technical-reference/automatically-created-managed-properties-in-sharepoint
                        }*/

                        // https://docs.microsoft.com/en-us/sharepoint/dev/general-development/fast-query-language-fql-syntax-reference#fql_range_operator
                        if (this.moment(refinementToken, this.moment.ISO_8601, true).isValid()) {

                            if (filterValue.operator === FilterComparisonOperator.Gt || filterValue.operator === FilterComparisonOperator.Geq) {
                                refinementToken = `range(${refinementToken},max)`;
                            }

                            // Ex: scenario ('older than a year')
                            if (filterValue.operator === FilterComparisonOperator.Leq || filterValue.operator === FilterComparisonOperator.Lt) {
                                refinementToken = `range(min,${refinementToken})`;
                            }
                        }

                        refinementQueryConditions.push(`${filter.filterName}:${refinementToken}`);
                    }
                }
            }
        });

        return refinementQueryConditions;
    }

    /**
     * Ensures the result source id value is a valid GUID
     * @param value the result source id
     */
    private validateSourceId(value: string): string {
        if (value.length > 0) {
            if (!(/^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/).test(value)) {
                return "The provided value is not a valid GUID";
            }
        }

        return '';
    }
}