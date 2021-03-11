import { IDataSource, IServiceKeysConfiguration } from "./IDataSource";
import { IDataSourceData } from "./IDataSourceData";
import { IPropertyPaneGroup } from "@microsoft/sp-property-pane";
import { ServiceScope } from '@microsoft/sp-core-library';
import { IDataContext } from './IDataContext';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { FilterBehavior } from "./FilterBehavior";
import { ITemplateSlot } from "../../common/ITemplateSlot";
import { IDataFilter } from "../search/IDataFilter";
import { IComboBoxOption } from "office-ui-fabric-react";

export abstract class BaseDataSource<T> implements IDataSource {

    protected serviceScope: ServiceScope;

    protected _properties: T;
    private _serviceKeys: IServiceKeysConfiguration;
    private _context: WebPartContext;
    private _render: () => void | Promise<void>;

    get properties(): T {
        return this._properties;
    }

    set properties(properties: T) {
        this._properties = properties;
    }

    get render(): () => void | Promise<void> {
        return this._render;
    }

    set render(renderFunc: () => void | Promise<void>) {
        this._render = renderFunc;
    }

    get context(): WebPartContext {
        return this._context;
    }

    set context(context: WebPartContext) {
        this._context = context;
    }

    get serviceKeys(): IServiceKeysConfiguration {
        return this._serviceKeys;
    }

    set serviceKeys(keys: IServiceKeysConfiguration) {
        this._serviceKeys = keys;
    }

    public constructor(serviceScope: ServiceScope) {
        this.serviceScope = serviceScope;
    }
    public abstract getAvailableProperties(): Promise<IComboBoxOption[]>;

    public onInit(): void | Promise<void> {
    }

    public abstract getData(dataContext?: IDataContext): Promise<IDataSourceData>;

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

        // Returns an empty configuration by default
        return [];
    }

    public getFilterBehavior(): FilterBehavior {

        // Filtering capabilioty by default
        return FilterBehavior.Static;
    }

    public getAppliedFilters(): IDataFilter[] {
        return [];
    }

    public abstract getItemCount(): number;

    public getTemplateSlots(): ITemplateSlot[] {

        // No slots by default
        return [];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        // Do nothing by default      
    }
}