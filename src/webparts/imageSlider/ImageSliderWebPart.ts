import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";
import * as strings from "ImageSliderWebPartStrings";
import ImageSlider from "./components/ImageSlider";
import { IImageSliderProps, IImageItem } from "./components/IImageSliderProps";
import {
  PropertyFieldCustomList,
  CustomListFieldType,
} from "sp-client-custom-fields/lib/PropertyFieldCustomList";
import { PropertyFieldAlignPicker } from "sp-client-custom-fields/lib/PropertyFieldAlignPicker";

export interface IImageSliderWebPartProps {
  items: Array<IImageItem>;
  bottomIndicators: boolean;
  prevNextControls: boolean;
  autoPlay: boolean;
  autoPlayInterval: number;
  width: string;
  height: string;
  captionPosition: string;
}

export default class ImageSliderWebPart extends BaseClientSideWebPart<IImageSliderWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IImageSliderProps> = React.createElement(
      ImageSlider,
      {
        items: this.properties.items,
        bottomIndicators: this.properties.bottomIndicators,
        prevNextControls: this.properties.prevNextControls,
        autoPlay: this.properties.autoPlay,
        autoPlayInterval: this.properties.autoPlayInterval,
        width: this.properties.width,
        height: this.properties.height,
        captionPosition: this.properties.captionPosition,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  //@ts-ignore
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyFieldCustomList("items", {
                  label: strings.DataFieldLabel,
                  value: this.properties.items,
                  headerText: "Manage Items",
                  fields: [
                    {
                      id: "header",
                      title: "Header",
                      required: true,
                      type: CustomListFieldType.string,
                    },
                    {
                      id: "description",
                      title: "Description",
                      required: true,
                      type: CustomListFieldType.string,
                    },
                    {
                      id: "enabled",
                      title: "Enabled",
                      required: true,
                      type: CustomListFieldType.boolean,
                    },
                    {
                      id: "picture",
                      title: "Picture",
                      required: true,
                      type: CustomListFieldType.string,
                    },
                    {
                      id: "linkUrl",
                      title: "Link Url",
                      required: false,
                      type: CustomListFieldType.string,
                      hidden: true,
                    },
                    {
                      id: "linkText",
                      title: "Link Text",
                      required: false,
                      type: CustomListFieldType.string,
                      hidden: true,
                    },
                  ],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  context: this.context,
                  properties: this.properties,
                  key: "ImageSliderCustomListField",
                }),
              ],
            },
            {
              groupName: strings.ControlsSettingsGroupName,
              groupFields: [
                PropertyPaneToggle("bottomIndicators", {
                  label: strings.BottomIndicatorsFieldLabel,
                }),
                PropertyPaneToggle("prevNextControls", {
                  label: strings.PrevNextControlsFieldLabel,
                }),
              ],
            },
            {
              groupName: strings.AutoPlaySettingsGroupName,
              groupFields: [
                PropertyPaneToggle("autoPlay", {
                  label: strings.AutoPlayFieldLabel,
                }),
                PropertyPaneSlider("autoPlayInterval", {
                  label: strings.AutoPlayIntervalFieldLabel,
                  min: 1,
                  max: 10,
                  step: 1,
                }),
              ],
            },
            {
              groupName: strings.LayoutSettingsGroupName,
              groupFields: [
                PropertyPaneTextField("width", {
                  label: strings.WidthFieldLabel,
                }),
                PropertyPaneTextField("height", {
                  label: strings.HeightFieldLabel,
                }),
                PropertyFieldAlignPicker("captionPosition", {
                  label: strings.CaptionPositionFieldLabel,
                  initialValue: this.properties.captionPosition,
                  onPropertyChanged: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  properties: this.properties,
                  key: "ImageSliderAlignPickerField",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
