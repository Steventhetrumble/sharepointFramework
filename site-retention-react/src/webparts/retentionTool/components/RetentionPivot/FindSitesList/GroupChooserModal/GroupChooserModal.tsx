import * as React from "react";
import { Modal, IDragOptions } from "office-ui-fabric-react/lib/Modal";
import { ISPSite } from "../../../common/IObjects";
import {
  DefaultButton,
  PrimaryButton
} from "office-ui-fabric-react/lib/Button";
import { getId } from "office-ui-fabric-react/lib/Utilities";
import { ContextualMenu } from "office-ui-fabric-react/lib/ContextualMenu";
import {
  modalContainer,
  modalBody,
  modalComponents,
  modalStackTokens
} from "./GroupChooserModal.styles";
import {
  Stack,
  Text,
  Sticky,
  StickyPositionType
} from "office-ui-fabric-react";
import {
  stackStyles,
  stackItemStyles,
  title,
  itemAlignmentsStackTokens,
  container
} from "../../../RetentionTool.styles";
import { UserNamesList } from "./UserNameList/UserNameList";
import { GroupDropdown } from "./GroupDropdown/GroupDropdown";
import {
  IGroupChooserModalProps,
  IGroupChooserModalState
} from "./GroupChooserModal.types";

export class GroupChooserModal extends React.Component<
  IGroupChooserModalProps,
  IGroupChooserModalState
> {
  constructor(props: IGroupChooserModalProps) {
    super(props);
    console.log("modal");
    console.log(this.props.groups);
  }
  // Use getId() to ensure that the IDs are unique on the page.
  // (It's also okay to use plain strings without getId() and manually ensure uniqueness.)
  private _titleId: string = getId("title");
  private _subtitleId: string = getId("subText");

  public render(): JSX.Element {
    return (
      <Modal
        titleAriaId={this._titleId}
        subtitleAriaId={this._subtitleId}
        isOpen={this.props.showModal}
        onDismiss={this.props.onClickCancel}
        isBlocking={false}
      >
        <Stack styles={modalContainer}>
          <Sticky stickyPosition={StickyPositionType.Header}>
            <Stack styles={stackStyles} tokens={itemAlignmentsStackTokens}>
              <Stack.Item align="stretch" styles={stackItemStyles}>
                <Text styles={title}>Title</Text>
              </Stack.Item>
            </Stack>
          </Sticky>

          <Stack verticalAlign="center" styles={modalBody}>
            <Stack
              horizontal
              horizontalAlign="space-evenly"
              verticalAlign="start"
            >
              <Stack.Item styles={modalComponents}>
                <GroupDropdown
                  site={this.props.siteSelected}
                  groups={this.props.groups}
                  provider={this.props.provider}
                  displayOptions={this.props.displayOptions}
                  onChangeSelectedIds={this.props.onChangeSelectedIds}
                />
              </Stack.Item>
              <Stack.Item styles={modalComponents}>
                <UserNamesList users={this.props.users} />
              </Stack.Item>
            </Stack>
          </Stack>

          <Sticky stickyPosition={StickyPositionType.Footer}>
            <Stack horizontal horizontalAlign="end" verticalAlign="end">
              <Stack.Item>
                <PrimaryButton
                  onClick={() => {
                    this.props.onClickAddTarget();
                  }}
                  style={{ marginRight: "8px" }}
                >
                  Add Site
                </PrimaryButton>
              </Stack.Item>
              <Stack.Item>
                <DefaultButton onClick={this.props.onClickCancel}>
                  Cancel
                </DefaultButton>
              </Stack.Item>
            </Stack>
          </Sticky>
        </Stack>
      </Modal>
    );
  }
}
