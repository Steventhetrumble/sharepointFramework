import * as React from "react";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import { Stack } from "office-ui-fabric-react";
import {
  SPGroupList,
  ISPUsers,
  SPGroup,
  ISPSite
} from "../../../../common/IObjects";
import { ISharepointRestProvider } from "../../../../dataproviders/ISharepointRestProvider.types";

export interface IGroupDropdownProps {
  groups: SPGroupList;
  provider: ISharepointRestProvider;
  site: ISPSite;
  displayOptions: IDropdownOption[];
  onChangeSelectedIds(ids: number[]): void;
}

export interface IGroupDropdownState {
  selectedItems: string[];
  selectedGroupIds: number[];
}

export class GroupDropdown extends React.Component<
  IGroupDropdownProps,
  IGroupDropdownState
> {
  constructor(props: IGroupDropdownProps) {
    super(props);

    this.state = {
      selectedItems: [],
      selectedGroupIds: []
    };
  }

  public componentDidUpdate(
    previousProps: IGroupDropdownProps,
    previousState: IGroupDropdownState
  ) {
    if (previousState.selectedGroupIds !== this.state.selectedGroupIds) {
      this.props.onChangeSelectedIds(this.state.selectedGroupIds);
    }
  }

  public render() {
    const { selectedItems } = this.state;

    return (
      <Stack verticalAlign="start">
        <Dropdown
          placeholder="Select options"
          label="Multi-select controlled example"
          selectedKeys={selectedItems}
          onChange={this._onChange}
          multiSelect
          options={this.props.displayOptions}
          styles={{ dropdown: { width: "20vw" } }}
        />
      </Stack>
    );
  }

  private _onChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    const newSelectedItems = [...this.state.selectedItems];
    const newSelectedIds = [...this.state.selectedGroupIds];
    if (item.selected) {
      // add the option if it's checked
      newSelectedItems.push(item.key as string);
      newSelectedIds.push(Number(item.key));
    } else {
      // remove the option if it's unchecked
      const currIndex = newSelectedItems.indexOf(item.key as string);
      const numIndex = newSelectedIds.indexOf(Number(item.key));

      if (numIndex > -1) {
        newSelectedIds.splice(numIndex, 1);
      }
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
    }
    this.setState({
      selectedItems: newSelectedItems,
      selectedGroupIds: newSelectedIds
    });
  };
}
