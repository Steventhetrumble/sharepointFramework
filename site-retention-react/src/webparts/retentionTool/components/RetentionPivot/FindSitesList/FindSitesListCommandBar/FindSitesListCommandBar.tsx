import * as React from 'react';
import { IFindSitesListCommandBarProps, IFindSitesListCommandBarState} from './FindSitesListCommandBar.types';
import { Stack, List, CommandBar, DirectionalHint, IButtonProps, CommandBarButton } from 'office-ui-fabric-react';

export class FindSitesListCommandBar extends React.Component<IFindSitesListCommandBarProps,IFindSitesListCommandBarState>{
    constructor(props: IFindSitesListCommandBarProps){
        super(props);

    }
    public render(): JSX.Element {
        const customButton = (props: IButtonProps) => {
            return (
                <CommandBarButton
                    {...props}
                    styles={{
                        ...props.styles,
                        textContainer: { fontSize: 18 },
                        icon: { color: '#E20000' }
                    }}
                />
            );
        };
        
        return(
            <CommandBar
            buttonAs={customButton}
            items={this.getItems()}
            ariaLabel={'Use left and right arrow keys to navigate between commands'}
            
            />
        );
    }
    private getItems = () => {
        return [
            {
                key: 'Inspect',
                name: 'Inspect',
                cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
                iconProps: {
                    iconName: 'OpenInNewWindow'
                },
                ariaLabel: 'Inspect',
                onClick: () => { this.props.onClickInspect();}
            }
        ];
    }
}