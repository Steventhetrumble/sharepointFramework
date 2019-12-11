import * as React from "react";
import {
  Stack,
  Icon,
  IIconProps,
  IconButton,
  DefaultPalette,
  FontSizes
} from "office-ui-fabric-react";
import { arrowIconStyle } from "./Carousel.styles";

// initializeIcons();
const leftChevron: IIconProps = {
  iconName: "ChevronLeft",
  styles: {
    root: {
      fontSize: FontSizes.superLarge
    }
  }
};

export const LeftArrow = ({ goToPrevSlide }) => {
  return (
    <Stack verticalAlign="center">
      <IconButton
        onClick={goToPrevSlide}
        iconProps={leftChevron}
        title="left"
        ariaLabel="left"
        styles={arrowIconStyle}
      />
    </Stack>
  );
};
