import * as React from 'react';
import styles from './DocumentCardExample.module.scss';
import { IDocumentCardExampleProps } from './IDocumentCardExampleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  
  List
} from 'office-ui-fabric-react/lib/List';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import * as ExampleList from './List/List.Basic.Example';

export default class DocumentCardExample extends React.Component<IDocumentCardExampleProps, {}> {
  constructor(props: IDocumentCardExampleProps) {
    super(props);
  }

  public componentDidMount(){
    console.log("component did mount");
  }

  public render(): JSX.Element {
    const {} = this.props;
  
    return (
      
      <div className={ styles.documentCardExample}>  
      <div className={styles.container}>  
        <div className={styles.row} >  
          <div className={styles.subTitle}>  
         
          
          <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint Framework Development</span>  
          <p className="ms-font-l ms-fontColor-white">Demo : Retrieve Site Data from SharePoint Rest Api</p>  
          <p className="ms-font-l ms-fontColor-white" id="Title">Title of Root Site : </p>  
          <p className="ms-font-l ms-fontColor-white" id="Description">Description : </p>  
          <p className="ms-font-l ms-fontColor-white" id="Url">Url : </p>  
          <p className="ms-font-l ms-fontColor-white" id="Size">Size :</p>  
          <p className="ms-font-l ms-fontColor-white" id="Results">Results: </p>  
          
          </div>  
        </div>
        <div className={styles.row}>  
        <div className={styles.subTitle}>Site Details</div>  
        <div className={styles.row}>
          <div id="SortButtons"><PrimaryButton  disabled >Sort by Last Item mod</PrimaryButton><PrimaryButton disabled>Sort by least Recent Views</PrimaryButton></div>
          <div ><PrimaryButton disabled><span>left</span></PrimaryButton><PrimaryButton  >right</PrimaryButton></div>
        </div>
        <div id="spListContainer" /> 
        <ExampleList.ListBasicExample></ExampleList.ListBasicExample>
        
        </div>  
      </div>  
    </div>
    );
  }
}
