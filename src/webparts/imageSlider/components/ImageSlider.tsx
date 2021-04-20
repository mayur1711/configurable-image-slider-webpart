import * as React from "react";
import styles from "./ImageSlider.module.scss";
import { IImageSliderProps } from "./IImageSliderProps";
import Carousel from "react-bootstrap/Carousel";

export default class ImageSlider extends React.Component<
  IImageSliderProps,
  {}
> {
  public componentDidMount() {
    if (document.getElementsByClassName("carousel-control left").length > 0) {
      document
        .getElementsByClassName("carousel-control left")[0]
        .setAttribute("class", "carousel-control-prev");
    }
    if (document.getElementsByClassName("carousel-control right").length > 0) {
      document
        .getElementsByClassName("carousel-control right")[0]
        .setAttribute("class", "carousel-control-next");
    }
  }

  public render(): React.ReactElement<IImageSliderProps> {
    const {
      items,
      bottomIndicators,
      prevNextControls,
      autoPlay,
      autoPlayInterval,
      width,
      height,
      captionPosition,
    } = this.props;

    return (
      <div className={styles.imageSlider} style={{ width: width }}>
        <Carousel
          slide={autoPlay}
          interval={autoPlayInterval * 1000}
          indicators={bottomIndicators}
          controls={prevNextControls}
          prevIcon={
            <span
              className="carousel-control-prev-icon"
              aria-hidden="true"
            ></span>
          }
          nextIcon={
            <span
              className="carousel-control-next-icon"
              aria-hidden="true"
            ></span>
          }
        >
          {items
            .filter((item) => item.enabled.toString() == "true")
            .map((item, key) => {
              return (
                <Carousel.Item key={key} index={key} className="carousel-item">
                  <img
                    height={height}
                    className="d-block w-100"
                    src={item.picture}
                    alt={item.header}
                  />
                  <Carousel.Caption
                    className={
                      captionPosition == "center"
                        ? ""
                        : captionPosition == "left"
                        ? "left-align"
                        : "right-align"
                    }
                  >
                    <div>
                      <h3>{item.header}</h3>
                      <p>{item.description}</p>
                      <HyperLinkControl
                        linkUrl={item.linkUrl}
                        linkText={item.linkText}
                      ></HyperLinkControl>
                    </div>
                  </Carousel.Caption>
                </Carousel.Item>
              );
            })}
        </Carousel>
      </div>
    );
  }
}

export const HyperLinkControl = (props: any) => {
  if (props.linkUrl && props.linkText) {
    return (
      <a href={props.linkUrl} className="btn btn-sm btn-primary">
        {props.linkText}
      </a>
    );
  }
  return null;
};
