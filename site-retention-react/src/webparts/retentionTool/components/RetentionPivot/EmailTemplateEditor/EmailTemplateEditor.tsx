import * as React from "react";
import {
  IEmailTemplateEditorProps,
  IEmailTemplateEditorState
} from "./EmailTemplateEditor.types";
import {
  Stack,
  List,
  Sticky,
  Label,
  StickyPositionType,
  PrimaryButton
} from "office-ui-fabric-react";
import { row } from "../../RetentionTool.styles";
import { EmailTemplateEditorTextBox } from "./EmailTemplateEditorTextBox/EmailTemplateEditorTextBox";
import { pivotStackItemStyles } from "../../RetentionTool.styles";
export class EmailTemplateEditor extends React.Component<
  IEmailTemplateEditorProps,
  IEmailTemplateEditorState
> {
  constructor(props: IEmailTemplateEditorProps) {
    super(props);

    this.state = {
      totalRows: 0
    };
  }
  public render(): JSX.Element {
    return (
      <Sticky stickyPosition={StickyPositionType.Header}>
        <Stack styles={row}>
          <Stack.Item align="stretch" styles={pivotStackItemStyles}>
            <EmailTemplateEditorTextBox
              email={this.props.emails[this.props.email]}
            ></EmailTemplateEditorTextBox>
          </Stack.Item>
          <Stack.Item align="end" styles={pivotStackItemStyles}>
            <PrimaryButton onClick={this._saveEmail}>hey</PrimaryButton>
          </Stack.Item>
        </Stack>
      </Sticky>
    );
  }

  private _saveEmail = () => {
    console.log("this email saved--- not!");
  };
}
