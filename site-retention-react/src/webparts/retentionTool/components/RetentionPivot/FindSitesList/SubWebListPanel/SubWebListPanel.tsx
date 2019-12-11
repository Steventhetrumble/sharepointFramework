import { DefaultButton, PrimaryButton } from "office-ui-fabric-react/lib/Button"
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel"
import * as React from "react"
import {
  ISubWebListPanelProps,
  ISubWebListPanelState
} from "./SubWebListPanel.types"
import {
  Stack,
  Label,
  MarqueeSelection,
  DetailsList,
  SelectionMode,
  DetailsListLayoutMode,
  IColumn,
  Text
} from "office-ui-fabric-react"
import { ISPSite } from "../../../common/IObjects"
import {
  stackStyles,
  itemAlignmentsStackTokens,
  stackItemStyles,
  title
} from "../../../RetentionTool.styles"

export class SubWebListPanel extends React.Component<
  ISubWebListPanelProps,
  ISubWebListPanelState
> {
  private _columns: IColumn[]

  constructor(props: ISubWebListPanelProps) {
    super(props)

    this._columns = [
      {
        key: "column1",
        name: "Title",
        fieldName: "title",
        minWidth: 40,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: ISPSite) => {
          return <span>{item.Title}</span>
        }
      },
      {
        key: "column1",
        name: "Relative Url",
        fieldName: "url",
        minWidth: 60,
        maxWidth: 300,
        isResizable: true,
        onRender: (item: ISPSite) => {
          return <span>{item.Url}</span>
        }
      },
      {
        key: "column2",
        name: "Views Recent",
        fieldName: "viewsrecent",
        minWidth: 60,
        maxWidth: 100,
        isResizable: true,
        onRender: (item: ISPSite) => {
          return <span>{item.ViewsRecent}</span>
        }
      },
      {
        key: "column3",
        name: "Date User Last Modified",
        fieldName: "dateuserlastmodified",
        minWidth: 120,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: ISPSite) => {
          return <span>{item.LastItemUserModifiedDateFomatted}</span>
        }
      }
    ]
  }

  public render() {
    return (
      <Stack>
        <Panel
          isOpen={this.props.showPanel}
          closeButtonAriaLabel="Close"
          onDismiss={this.props.onClickClose}
          type={PanelType.large}
          headerText="Large Panel"
          onRenderFooterContent={this._onRenderFooterContent}
          onRenderHeader={this._onRenderHeader}
          onRenderBody={this._onRenderBody}
        />
      </Stack>
    )
  }

  private _onRenderFooterContent = () => {
    return (
      <Stack horizontal>
        <Stack.Item>
          <PrimaryButton
            onClick={() => {
              this.props.onClickAddSite()
            }}
            style={{ marginRight: "8px" }}
          >
            Add Site
          </PrimaryButton>
        </Stack.Item>
        <Stack.Item>
          <DefaultButton onClick={this.props.onClickClose}>
            Cancel
          </DefaultButton>
        </Stack.Item>
      </Stack>
    )
  }

  private _onRenderHeader = () => {
    return (
      <Stack styles={stackStyles} tokens={itemAlignmentsStackTokens}>
        <Stack.Item align="start" styles={stackItemStyles}>
          <Text styles={title}>Deletion Process</Text>
        </Stack.Item>
      </Stack>
    )
  }
  private _onRenderBody = () => {
    if (this.props.subSites != null && this.props.subSites.length > 0) {
      return (
        <DetailsList
          compact={true}
          items={this.props.subSites}
          columns={this._columns}
          selectionMode={SelectionMode.none}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionPreservedOnEmptyClick={true}
          checkButtonAriaLabel="Row checkbox"
        />
      )
    } else {
      return <Label disabled={true}>{"No Sub Webs Present"}</Label>
    }
  }
}
