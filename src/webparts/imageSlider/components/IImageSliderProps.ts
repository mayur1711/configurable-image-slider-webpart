export interface IImageSliderProps {
  items: Array<IImageItem>;
  bottomIndicators: boolean;
  prevNextControls: boolean;
  autoPlay: boolean;
  autoPlayInterval: number;
  width: string;
  height: string;
  captionPosition: string;
}

export interface IImageItem {
  header: string;
  description: string;
  enabled: boolean;
  picture: string;
  linkUrl: string;
  linkText: string;
}
