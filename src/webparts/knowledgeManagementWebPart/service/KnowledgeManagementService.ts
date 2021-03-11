import ISearchResult from "../model/ISearchResultData";
import { ISearchQuery, SearchResults } from '@pnp/sp/search';
import { sp } from "@pnp/sp";
import { IDocument } from "../model/IDocument";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export default class KnowledgeManagementService {

    public static LoadData(_rowLimit: number, _queryText: string) {
        return new Promise<ISearchResult>((resolve, reject) => {
           /* sp.search(<ISearchQuery> {
                Querytext: _queryText,
                RowLimit: _rowLimit,
                SelectProperties: ["CreatedBy", "ModifiedBy", "Title", "FileType", "ModifiedOWSDATE", "Size"]
            }).then((result) => {
                let data: ISearchResult = {
                    PrimaryQueryResult: result.PrimarySearchResults,
                    SecondaryQueryResults: result.RawSearchResults.SecondaryQueryResults,
                    Refiner: result.RawSearchResults.PrimaryQueryResult.RefinementResults ? result.RawSearchResults.PrimaryQueryResult.RefinementResults.Refiners : null       
                };
                resolve(data);
            })
            .catch((error) => {
                //TODO Error
                reject("error");
            });*/
        });
    }

    public static GenerateDocuments(_items: ISearchResult) {
        const items: IDocument[] = [];
        let countItem = 0;

        _items.PrimaryQueryResult.forEach((item, index) => {
            items.push({
                key: countItem.toString(),
                name: item.Title,
                value: item.Title,
                iconName: item.FileType,
                fileType: item.FileType,
                modifiedBy: item["ModifiedBy"],
                dateModified: item["ModifiedOWSDATE"],
                dateModifiedValue: 0,
                fileSize: item["Size"] ? item["Size"].toString() : null,
                fileSizeRaw: item["Size"] ? item["Size"] : null  
            });

            countItem++;
        });
        return items;
    }
}