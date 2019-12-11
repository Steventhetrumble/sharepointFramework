import * as React from 'react';
import styles from './ManageLists.module.scss';
import { IManageListsProps } from './IManageListsProps';
import IManageListsState from './ImanageListsState';
import SPTitleBlock from './TitleBlock/SPTitleBlock';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { SPPivot } from './SPPivot/ISPPivot';


export default class ManageLists extends React.Component<IManageListsProps, IManageListsState> {
  constructor(props: IManageListsProps) {
    super(props);
    this.state = {
      lists: [],
      rootSite: this.props.site
    };
  }

  public render(): React.ReactElement<IManageListsProps> {
    return (
      
        <Fabric>
          <div className={styles.manageLists}>
            <div className={styles.container}>
            
            
            <SPTitleBlock site={this.state.rootSite} spHttpClient={this.props.spHttpClient}></SPTitleBlock>
            <SPPivot site={this.state.rootSite} spHttpClient={this.props.spHttpClient} provider={this.props.provider}></SPPivot>
            </div>
          </div>
        </Fabric>
    );
  }
}
