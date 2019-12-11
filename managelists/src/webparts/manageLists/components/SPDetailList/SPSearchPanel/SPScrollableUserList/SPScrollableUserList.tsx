import * as React from 'react';
import { getTheme, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react/lib/ScrollablePane';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { ISPGroupEmails, ISPGroupEmail } from '../../../../common/IObjects';

const theme = getTheme();
const classNames = mergeStyleSets({
  wrapper: {
    height: '40vh',
    position: 'relative',
    maxHeight: 'inherit',
    width: '60vh'
  },
  
});

export interface ISPScrollableUserListProps{
    emailGroups: ISPGroupEmails;
}

export interface ISPScrollableUserListState{

}



export class SPScrollableUserList extends React.Component<ISPScrollableUserListProps, ISPScrollableUserListState> {

  constructor(props: ISPScrollableUserListProps) {
    super(props);

    // Using splice prevents the colors from being duplicated
    
  }

  public render() {
    let contentAreas = this.props.emailGroups.groups.map(this._createContentArea);

    return (
      <div className={classNames.wrapper}>
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto} >
        
            {...contentAreas}
            
        </ScrollablePane>
      </div>
    );
  }

  private _createContentArea = (item: ISPGroupEmail) => {
      let emailList: string;
      let count: number = 0;
      item.userEmails.forEach((s: string) => {
        emailList += `;${s} `;
          count++;
      });
      return (
          <div>
          
            <Text variant={'large'} block>
            {item.GroupName}
            </Text>
            <Text>{emailList}</Text>
          
          </div>

      );
  }
}
 