import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './KnowledgeManagementWebPart.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { msGraphBusiness } from '../../../business/msGraphBusiness';

const KnowledgeManagementWebPart = (props : { spfxContext: WebPartContext}) => {

  useEffect(() => {
    var toto = msGraphBusiness.GetCurrentUser(props.spfxContext);
    var docs   = msGraphBusiness.GetAllDocuments(props.spfxContext);
    
    console.log('hello');
  }, []);

    return (
      <div className={ styles.knowledgeManagementWebPart }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>

              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
};

export default KnowledgeManagementWebPart;
