import * as React from "react";
import { Stack, hiddenContentStyle, Sticky } from "office-ui-fabric-react";
import { RightArrow } from "./RightArrow";
import { LeftArrow } from "./LeftArrow";
import { Slide } from "./Slide";
import { landingPageCarousel } from "./Carousel.styles";

const images = [
  {
    original: "https://picsum.photos/id/1018/1000/600/",
    thumbnail: "https://picsum.photos/id/1018/250/150/"
  },
  {
    original: "https://picsum.photos/id/1015/1000/600/",
    thumbnail: "https://picsum.photos/id/1015/250/150/"
  },
  {
    original: "https://picsum.photos/id/1019/1000/600/",
    thumbnail: "https://picsum.photos/id/1019/250/150/"
  }
];

export interface ICarouselState {
  images: string[];
  currentIndex: number;
  translateValue: number;
}

export interface ICarouselProps {}

export class Carousel extends React.Component<ICarouselProps, ICarouselState> {
  private _initialOffset: number = -215;

  constructor(props) {
    super(props);

    this.state = {
      images: [
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/aurora.jpg",
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/canyon.jpg",
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/city.jpg",
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/desert.jpg",
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/mountains.jpg",
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/redsky.jpg",
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/sandy-shores.jpg",
        "https://s3.us-east-2.amazonaws.com/dzuz14/thumbnails/tree-of-life.jpg"
      ],
      currentIndex: 0,
      translateValue: this._initialOffset
    };
  }

  public goToPrevSlide = () => {
    if (this.state.translateValue === this._initialOffset) {
      return this.setState({
        currentIndex: this.state.images.length - 1,
        translateValue:
          (this.slideWidth() + 20) * -(this.state.images.length - 1) +
          this._initialOffset
      });
    }
    this.setState(prevState => ({
      currentIndex: prevState.currentIndex - 1,
      translateValue: prevState.translateValue + this.slideWidth() + 20
    }));
  };

  public goToNextSlide = () => {
    if (this.state.currentIndex === this.state.images.length - 1) {
      return this.setState({
        currentIndex: 0,
        translateValue: this._initialOffset
      });
    }
    this.setState(prevState => ({
      currentIndex: prevState.currentIndex + 1,
      translateValue: prevState.translateValue + -(this.slideWidth() + 20)
    }));
  };

  public slideWidth = () => {
    return document.querySelector(".slide").clientWidth;
  };

  public render(): JSX.Element {
    return (
      <Stack
        horizontalAlign="space-between"
        horizontal
        styles={{
          root: {
            overflow: "hidden",
            width: "62vw",
            maxHeight: "50vh",
            maxWidth: "600px",
            minWidth: "380px"
          }
        }}
      >
        <LeftArrow goToPrevSlide={this.goToPrevSlide} />

        <Stack
          horizontal
          tokens={{ childrenGap: 20 }}
          styles={{
            root: {
              width: "20vw",
              overflow: "visible",
              transform: `translateX(${this.state.translateValue}px)`,
              transition: "transform ease-out 0.45s"
            }
          }}
        >
          {this.state.images.map((image, i) => (
            <Stack.Item>
              <Slide key={i} image={image} />
            </Stack.Item>
          ))}
        </Stack>
        <RightArrow goToNextSlide={this.goToNextSlide} />
      </Stack>
    );
  }
}
