import * as React from 'react';
import styles from './KnowledgeManagementWebPart.module.scss';
import KMSearchBar from './KMSearchBar';
import { useState } from 'react';
import KMSearchResult from './KMSearchResult';
import { WebPartTitle } from '@pnp/spfx-controls-react';

const KnowledgeManagementWebPart = (props) => {

  const [searchBoxValue, setSearchBoxValue] = useState("");

    return (
      <div className={ styles.knowledgeManagementWebPart }>
        <WebPartTitle
                displayMode={props.webPartTitleProps.displayMode}
                title={props.webPartTitleProps.title}
                updateProperty={props.webPartTitleProps.updateProperty}
            />
        {
          props.ShowHideSearchBar &&
          <KMSearchBar 
              searchBoxValueProp={searchBoxValue}
              setSearchBoxValueProp={setSearchBoxValue}
               />
        }
          <KMSearchResult 
              ThemeView={props.Theme}
          />
      </div>
    );
  
};

export default KnowledgeManagementWebPart;
