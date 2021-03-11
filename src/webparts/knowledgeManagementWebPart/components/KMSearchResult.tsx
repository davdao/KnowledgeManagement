import * as React from 'react';
import { useEffect } from 'react';
import { ThemeBackgrounds } from '../../../common/Enum';
import KnowledgeManagementService from '../service/KnowledgeManagementService';
import styles from './KMSearchResult.module.scss';
import { useState } from 'react';
import { ISearchResult } from '@pnp/sp/search';

const DocumentCardView = React.lazy(() => import('../../../views/DocumentCardView/DocumentCardLayout'));
const SimpleListView = React.lazy(() => import('../../../views/ListView/ListViewLayout'));

const KMSearchResult = (props) => {
    
    const [resultData, setResultData] = useState(null);

    useEffect(() => {
        KnowledgeManagementService.LoadData(10, "path:https://ddaodev.sharepoint.com/sites/DemoComponents/Shared%20Documents").then((result) => {
            setResultData(result);
        });
      }, []);
      
    return(
        <div>
            <React.Suspense fallback={<></>}>
                {
                    resultData &&
                        renderSearchResult(props.ThemeView, resultData)
                }
            </React.Suspense>
        </div>
    );


};

function renderSearchResult(_themeView, _resultData: ISearchResult)
{
    switch(_themeView)
    {
        case ThemeBackgrounds.List:
            return(<SimpleListView listData={_resultData} columnList={["FileType", "Title", "Description"]} />);
        break;

        case ThemeBackgrounds.Brick:
            return(<DocumentCardView />);
        break;

        default:
            return(<SimpleListView listData={_resultData} />);
        break;
    }
}

export default KMSearchResult;