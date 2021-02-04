import * as React from 'react';
import styles from './KnowledgeManagementWebPart.module.scss';
import { useState, useEffect } from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import KMSearchBar from './KMSearchBar';
import KMSearchResult from './KMSearchResult';

const [searchBoxValue, setSearchBoxValue] = useState("");

const KnowledgeManagementWebPart = () => {
/*
 <KMSearchBar 
          searchBoxValueProp={(searchBoxValue)}
          searchBoxOnChangeProp={(e) => setSearchBoxValue(e)} />

*/
  useEffect(() => {


    //Get all result
  }, []);

  return (
    <div className={ styles.knowledgeManagementWebPart }>
     

      <KMSearchResult />

      <div>
        {"oho"}
        {searchBoxValue}
      </div>
    </div>
  );

};

export default KnowledgeManagementWebPart;