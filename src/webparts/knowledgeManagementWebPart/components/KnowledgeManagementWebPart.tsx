import * as React from 'react';
import styles from './KnowledgeManagementWebPart.module.scss';
import { IKnowledgeManagementWebPartProps } from './IKnowledgeManagementWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SharePointDataSourceServices } from '../../../service/SharePointServices/SharePointDataSourceServices';

export default class KnowledgeManagementWebPart extends React.Component<IKnowledgeManagementWebPartProps, {}> {

  public componentDidMount() {
    /*const _sharePointDataSourceServices = this.props.serviceScope.consume(SharePointDataSourceServices.serviceKey);
  
    _sharePointDataSourceServices.getAllMetadataSearchProperties().then((result) => {
      console.log(result);
    });*/

    
  }
  
  public render(): React.ReactElement<IKnowledgeManagementWebPartProps> {
    return (
      <div className={ styles.knowledgeManagementWebPart }>
        {"Hello World"}
      </div>
    );
  }
}
