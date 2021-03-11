
import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { useState, useEffect } from 'react';
import KnowledgeManagementService from '../../webparts/knowledgeManagementWebPart/service/KnowledgeManagementService';

const ListViewLayout = (props) => {

    const [itemDataFormated, setItemDataFormated] = useState(null);
      
/*
    useEffect(() => {
        //props.listData
        setItemDataFormated(KnowledgeManagementService.GenerateDocuments(props.listData));
    }, [itemDataFormated]);*/
    return(<div>
                {
                    <DetailsList 
                    items={KnowledgeManagementService.GenerateDocuments(props.listData)}
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                    selectionMode={SelectionMode.none}
                    columns={GenerateColumns(props.columnList)}
                    />
                }             
            </div>);
};

    function GenerateColumns(_columns: string[]) {
        const columns: IColumn[] = [];
        let counterColumn = 0;
        _columns.forEach((column, index) => {
            columns.push({
                key: _columns + "_" + counterColumn,
                name: column,
                fieldName: "modifiedBy",
                minWidth: 210,
                maxWidth: 350,
                isRowHeader: true,
                isResizable: true,
                //onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
            });
            counterColumn++;
        });
        return columns;
    }

export default ListViewLayout;