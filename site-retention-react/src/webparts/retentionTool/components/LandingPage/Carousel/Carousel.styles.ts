import {
  IStackStyles,
  IIconStyles,
  DefaultPalette,
  FontSizes
} from "office-ui-fabric-react";

export const landingPageCarousel: IStackStyles = {
  root: {
    position: "relative",
    overflow: "hidden"
  }
};

export const arrowIconStyle: IIconStyles = {
  root: {
    color: DefaultPalette.magentaLight,
    fontSize: FontSizes.superLarge,
    height: "10vh",
    width: "10vh",
    zIndex: 999
  }
};
