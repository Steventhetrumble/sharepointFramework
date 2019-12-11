import * as React from "react";
import {
  IRetentionPivotProps,
  IRetentionPivotState
} from "./RetentionPivot.types";
import { Pivot, PivotItem, Label, Stack } from "office-ui-fabric-react";
import { labels, contentContainer } from "./RetentionPivot.style";
import {
  stackStyles,
  itemAlignmentsStackTokens,
  stackItemStyles,
  row
} from "../RetentionTool.styles";
import { FindSitesList } from "./FindSitesList/FindSitesList";
import { TargetedSites } from "./TargetedSites/TargetedSites";
import { EmailTemplateEditor } from "./EmailTemplateEditor/EmailTemplateEditor";
import { ISPDate, ISPSite, Row, ListResult } from "../common/IObjects";

export class RetentionPivot extends React.Component<
  IRetentionPivotProps,
  IRetentionPivotState
> {
  constructor(props: IRetentionPivotProps) {
    super(props);
    this.state = {
      totalRows: 0,
      items: [],
      targets: null,
      targetRows: 0,
      emails: ["", ""]
    };
    this._findSites();
    this._findTargetedSites();
  }

  public render(): JSX.Element {
    return (
      <Pivot>
        <PivotItem
          headerText="Search List"
          headerButtonProps={{
            "data-order": 1,
            "data-title": "Search List"
          }}
        >
          <Stack styles={contentContainer}>
            <FindSitesList
              provider={this.props.provider}
              targets={this.state.targets}
              totalRows={this.state.totalRows}
              items={this.state.items}
              onSiteAddition={this._findTargetedSites}
            />
          </Stack>
        </PivotItem>
        <PivotItem headerText="Targeted Sites">
          <Stack styles={contentContainer}>
            <TargetedSites
              targets={this.state.targets}
              totalRows={this.state.targetRows}
              provider={this.props.provider}
              onSiteDeletion={this._findTargetedSites}
            />
          </Stack>
        </PivotItem>
        <PivotItem headerText="1st Round Email">
          <Stack styles={contentContainer}>
            <EmailTemplateEditor
              email={0}
              emails={this.state.emails}
            ></EmailTemplateEditor>
          </Stack>
        </PivotItem>
        <PivotItem headerText="2nd Round Email">
          <Stack styles={contentContainer}>
            <EmailTemplateEditor
              email={1}
              emails={this.state.emails}
            ></EmailTemplateEditor>
          </Stack>
        </PivotItem>
        <PivotItem headerText="Developer">
          <Stack verticalAlign="center" styles={contentContainer}>
            <Stack styles={row}>
              <Stack styles={stackStyles} tokens={itemAlignmentsStackTokens}>
                <Stack.Item styles={stackItemStyles}>
                  <Label styles={labels}>name: Steven Trumble</Label>
                </Stack.Item>
                <Stack.Item styles={stackItemStyles}>
                  <Label styles={labels}>
                    email: StevenAndrew.Trumbl1@rci.rogers.com
                  </Label>
                </Stack.Item>
                <Stack.Item styles={stackItemStyles}>
                  <Label styles={labels}>
                    Alternate email: Steven.a.trumble@gmail.ca
                  </Label>
                </Stack.Item>
              </Stack>
            </Stack>
          </Stack>
        </PivotItem>
      </Pivot>
    );
  }

  private _getEmails = (): void => {
    this.props.provider.getEmails().then((emails: string[]) => {
      this.setState({
        emails: emails
      });
    });
  };
  private _findTargetedSites = (): void => {
    this.props.provider.getTargetedSitesListing().then((result: ListResult) => {
      this.setState({
        targets: result,
        targetRows: result.value.length
      });
    });
  };

  private _findSites = (): void => {
    this.props.provider.getTotalSites().then((totalRows: number) => {
      this.setState({
        totalRows: totalRows
      });
      for (var i = 0; i < Math.ceil(totalRows / 500); i++) {
        this.props.provider.getSitesListing(i * 500).then((rows: Row[]) => {
          rows.forEach((_row: Row) => {
            this.props.provider
              .getSPDate(_row)
              .then((responseDate: ISPDate) => {
                let tempSPSite: ISPSite = {
                  Title: _row.Cells[2].Value,
                  Url: _row.Cells[3].Value,
                  ViewsLifeTime: Number(_row.Cells[6].Value),
                  ViewsRecent: Number(_row.Cells[7].Value),
                  Size: Number(_row.Cells[8].Value),
                  SiteDescription: _row.Cells[9].Value,
                  LastItemUserModifiedDateSharepoint:
                    responseDate.error == null ? responseDate.value : "ok",
                  LastItemUserModifiedDate:
                    responseDate.error == null ? responseDate.date : new Date(),
                  LastItemUserModifiedDatevalue:
                    responseDate.error == null ? responseDate.datevalue : 0,
                  LastItemUserModifiedDateFomatted:
                    responseDate.error == null
                      ? responseDate.dateFormatted
                      : "",
                  renderTemplateId: _row.Cells[16].Value
                };
                let tempSites: ISPSite[] = this.state.items;
                this.setState({
                  items: [...tempSites, tempSPSite]
                });
              });
          });
        });
      }
    });
  };
}
