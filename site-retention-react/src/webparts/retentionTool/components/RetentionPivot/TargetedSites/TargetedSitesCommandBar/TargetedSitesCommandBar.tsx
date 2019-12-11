import * as React from 'react';
import { ITargetedSitesCommandBarProps, ITargetedSitesCommandBarState} from './TargetedSitesCommandBar.types';
import { Stack, List, CommandBar, DirectionalHint, IButtonProps, CommandBarButton } from 'office-ui-fabric-react';

export class TargetedSitesCommandBar extends React.Component<ITargetedSitesCommandBarProps,ITargetedSitesCommandBarState>{
    constructor(props: ITargetedSitesCommandBarProps){
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
            overflowButtonProps={{
                ariaLabel: 'More commands',
                menuProps: {
                    items: [], // Items must be passed for typesafety, but commandBar will determine items rendered in overflow
                    isBeakVisible: true,
                    beakWidth: 20,
                    gapSpace: 10,
                    directionalHint: DirectionalHint.topCenter
                }
            }}
            buttonAs={customButton}
            overflowItems={this.getOverflowItems()}
            items={this.getItems()}
            farItems={this.getFarItems()}
            ariaLabel={'Use left and right arrow keys to navigate between commands'}
            
            />
        );
    }
    private getItems = () => {
        return [
            {
                key: 'deleteItem',
                name: 'Delete',
                cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
                iconProps: {
                    iconName: 'Delete'
                },
                ariaLabel: 'Delete',
                onClick: () => { this.props.onClickDelete();}
            },
            {
                key: 'edit',
                name: 'Edit',
                iconProps: {
                    iconName: 'Edit'
                },
                onClick: () => {this.props.onClickEdit();}
            }

        ];
    }

    private getOverflowItems = () => {
        return [
            {
                key: 'visitSiteList',
                name: 'Site List',
                iconProps: {
                    iconName: 'Emoji'
                },
                href: 'https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/Lists/TargetSitesForDeletionSteven1234/AllItems.aspx',
                ['data-automation-id']: 'visitSiteList',
                onClick: () => console.log('hey')
            },
            {
                key: 'visitEmailList',
                name: 'Email List',
                iconProps: {
                    iconName: 'Emoji2'
                },
                href: 'https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/Lists/TargetedEmailTemplates/AllItems.aspx',
                ['data-automation-id']: 'visitEmailList'
            }
        ];
    }

    private getFarItems = () => {
        return [
            {
                key: 'mail',
                name: 'Generate Mail',
                ariaLabel: 'Generate Mail',
                subMenuProps: {
                    items: [
                        {
                            key: 'email1',
                            name: 'email 1',
                            iconProps: {
                                iconName: 'Mail'
                            },
                            onClick: () => {
                                console.log("clicked");



                            },
                            ['data-automation-id']: 'newEmailButton1'
                        },
                        {
                            key: 'email2',
                            name: 'email 2',
                            iconProps: {
                                iconName: 'Mail'
                            },
                            onClick: () => console.log('email2'),
                            ['data-automation-id']: 'newEmailButton2'
                        }
                    ]
                },
                iconProps: {
                    iconName: 'Mail'
                }
            }
        ];
    }

}