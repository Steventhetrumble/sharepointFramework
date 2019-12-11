import * as React from "react";
import {
  FocusZone,
  FocusZoneDirection
} from "office-ui-fabric-react/lib/FocusZone";
import { List } from "office-ui-fabric-react/lib/List";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";
import {
  ITheme,
  mergeStyleSets,
  getTheme,
  getFocusStyle
} from "office-ui-fabric-react/lib/Styling";
import { createListItems, IExampleItem } from "@uifabric/example-data";
import { ISPUsers, ISPUser } from "../../../../common/IObjects";

export interface IUserNamesListProps {
  items?: IExampleItem[];
  users: ISPUser[];
}

export interface IUserNamesListState {}

interface IUserNamesListClassObject {
  container: string;
  itemCell: string;
  itemImage: string;
  itemContent: string;
  itemName: string;
  itemIndex: string;
  chevron: string;
}

const theme: ITheme = getTheme();
const { palette, semanticColors, fonts } = theme;

const classNames: IUserNamesListClassObject = mergeStyleSets({
  container: {
    overflow: "auto",
    maxHeight: "30vh"
  },
  itemCell: [
    getFocusStyle(theme, { inset: -1 }),
    {
      minHeight: 54,
      padding: 10,
      boxSizing: "border-box",
      borderBottom: `1px solid ${semanticColors.bodyDivider}`,
      display: "flex",
      selectors: {
        "&:hover": { background: palette.neutralLight }
      }
    }
  ],
  itemImage: {
    flexShrink: 0
  },
  itemContent: {
    marginLeft: 10,
    overflow: "hidden",
    flexGrow: 1
  },
  itemName: [
    fonts.xLarge,
    {
      whiteSpace: "nowrap",
      overflow: "hidden",
      textOverflow: "ellipsis"
    }
  ],
  itemIndex: {
    fontSize: fonts.small.fontSize,
    color: palette.neutralTertiary,
    marginBottom: 10
  },
  chevron: {
    alignSelf: "center",
    marginLeft: 10,
    color: palette.neutralTertiary,
    fontSize: fonts.large.fontSize,
    flexShrink: 0
  }
});

export class UserNamesList extends React.Component<IUserNamesListProps> {
  private _items: ISPUser[];
  constructor(props: IUserNamesListProps) {
    super(props);
  }

  public render(): JSX.Element {
    return (
      <FocusZone direction={FocusZoneDirection.vertical}>
        <div className={classNames.container} data-is-scrollable={true}>
          <List items={this.props.users} onRenderCell={this._onRenderCell} />
        </div>
      </FocusZone>
    );
  }

  private _onRenderCell = (
    item: ISPUser,
    index: number,
    isScrolling: boolean
  ): JSX.Element => {
    if (item != null) {
      return (
        <div className={classNames.itemCell} data-is-focusable={true}>
          <div className={classNames.itemContent}>
            <div className={classNames.itemName}>{item.Email}</div>
            <div className={classNames.itemIndex}>{`Item ${index}`}</div>
          </div>
        </div>
      );
    } else {
      return <div>hello</div>;
    }
  };
}
