import * as React from 'react';
import styles from './KMSearchBar.module.scss';
import { SearchBox } from 'office-ui-fabric-react';
import * as strings from 'KnowledgeManagementWebPartWebPartStrings';
import { useState } from 'react';

const KMSearchBar = () => {

    const [searchBoxValue, setSearchBoxValue] = useState("");
    return(
        <div className={styles.KMSearchBoxWrapper}>
            <SearchBox
                placeholder={strings.SearchBox.DefaultPlacerHolder}
                value={searchBoxValue}
                autoComplete= "off"
                //onChange={ (value: string) => props.searchBoxOnChangeProp(value) }
                onSearch={ null }
                onClear={ () => setSearchBoxValue("") }
            />
        </div>
    );
};

export default KMSearchBar;