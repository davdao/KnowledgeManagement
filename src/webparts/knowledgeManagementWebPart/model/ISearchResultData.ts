import { IRefiner, IResultTableCollection, ISearchResult } from "@pnp/sp/search";

export interface ISearchResultData {
    PrimaryQueryResult?: ISearchResult[];
    SecondaryQueryResults?: IResultTableCollection;
    Refiner: IRefiner[];
}

export default ISearchResultData;