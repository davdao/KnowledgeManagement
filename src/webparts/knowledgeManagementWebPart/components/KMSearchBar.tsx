import * as React from 'react';
import styles from './KMSearchBar.module.scss';
import { SearchBox } from 'office-ui-fabric-react';
import * as strings from 'KnowledgeManagementWebPartWebPartStrings';

export interface IKMSearchBarProps {
    searchBoxValueProp: string;
    setSearchBoxValueProp:(e) => void;
}
  

const KMSearchBar = (props) => {

    return(
        <div className={styles.KMSearchBoxWrapper}>
            <SearchBox
                placeholder={strings.SearchBox.DefaultPlacerHolder}
                value={props.searchBoxValueProp}
                autoComplete= "off"
                onChange={ (value: string) => props.setSearchBoxValueProp(value) }
                onSearch={ null }
                onClear={ () => props.setSearchBoxValueProp("") }
            />
        </div>
    );
};

export default KMSearchBar;