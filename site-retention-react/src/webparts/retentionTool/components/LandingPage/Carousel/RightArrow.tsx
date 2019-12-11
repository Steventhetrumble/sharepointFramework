import * as React from "react";
import {
  Stack,
  Icon,
  IconButton,
  IIconProps,
  DefaultPalette,
  FontSizes
} from "office-ui-fabric-react";
import { arrowIconStyle } from "./Carousel.styles";

// initializeIcons();
const rightChevron: IIconProps = {
  iconName: "ChevronRight",
  styles: {
    root: {
      fontSize: FontSizes.superLarge
    }
  }
};

export const RightArrow = ({ goToNextSlide }) => {
  return (
    <Stack verticalAlign="center">
      <IconButton
        onClick={goToNextSlide}
        iconProps={rightChevron}
        title="right"
        ariaLabel="right"
        styles={arrowIconStyle}
      />
    </Stack>
  );
};
