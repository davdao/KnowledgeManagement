import * as React from 'react';
import styles from './KnowledgeManagementWebPart.module.scss';
import { IKnowledgeManagementWebPartProps } from './IKnowledgeManagementWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class KnowledgeManagementWebPart extends React.Component<IKnowledgeManagementWebPartProps, {}> {
  public render(): React.ReactElement<IKnowledgeManagementWebPartProps> {
    return (
      <div className={ styles.knowledgeManagementWebPart }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
