import * as React from 'react';
import styles from './KnowledgeManagementWebPart.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import KMSearchBar from './KMSearchBar';
import { useState } from 'react';

const KnowledgeManagementWebPart = () => {

  const [searchBoxValue, setSearchBoxValue] = useState("");
    return (
      <div className={ styles.knowledgeManagementWebPart }>
          <KMSearchBar 
              searchBoxValueProp={searchBoxValue}
              setSearchBoxValueProp={setSearchBoxValue} />
              <div>
                {searchBoxValue}
              </div>
      </div>
    );
  
}

export default KnowledgeManagementWebPart;
