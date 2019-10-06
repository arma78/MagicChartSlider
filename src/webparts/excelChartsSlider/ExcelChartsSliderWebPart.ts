import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPHttpClient} from "@microsoft/sp-http";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  PropertyPaneSlider,
  PropertyPaneToggle,
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'ExcelChartsSliderWebPartStrings';
import ExcelChartsSlider from './components/ExcelChartsSlider';
import { IExcelChartsSliderProps } from './components/IExcelChartsSliderProps';

export interface IExcelChartsSliderWebPartProps {

  description: string;
  listName: string;
  wbName: string;
  siteurl: string;
  spHttpClient: SPHttpClient;
  title: string;
  showThumbs:boolean;
  autoPlay:boolean;
  infiniteLoop:boolean;
  interval:number;
  showArrows:boolean;
  showStatus:boolean;
  swipeable:boolean;
  stopOnHover:boolean;
  showIndicators:boolean;
  transitionTime:number;
}

export default class ExcelChartsSliderWebPart extends BaseClientSideWebPart<IExcelChartsSliderWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExcelChartsSliderProps > = React.createElement(
      ExcelChartsSlider,
      {
        title:this.properties.title,
        description: this.properties.description,
        listName: this.properties.listName,
        wbName: this.properties.wbName,
        spHttpClient: this.context.spHttpClient,
        siteurl: this.context.pageContext.web.absoluteUrl,
        showThumbs: this.properties.showThumbs,
        autoPlay: this.properties.autoPlay,
        infiniteLoop: this.properties.infiniteLoop,
        interval: this.properties.interval,
        showArrows: this.properties.showArrows,
        showStatus: this.properties.showStatus,
        swipeable: this.properties.swipeable,
        stopOnHover: this.properties.stopOnHover,
        showIndicators: this.properties.showIndicators,
        transitionTime: this.properties.transitionTime,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleFieldLabel,
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneTextField('wbName', {
                  label: strings.wbNameFieldLabel
                }),
                PropertyPaneToggle("showThumbs", {
                  label: strings.showThumbsFieldLabel,
                  checked: true,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                 PropertyPaneToggle("autoPlay", {
                  label: strings.autoPlayFieldLabel,
                  checked: true,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                PropertyPaneToggle("showArrows", {
                  label: strings.showArrowsFieldLabel,
                  checked: true,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                 PropertyPaneToggle("showStatus", {
                  label: strings.showStatusFieldLabel,
                  checked: true,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                PropertyPaneToggle("stopOnHover", {
                  label: strings.stopOnHoverFieldLabel,
                  checked: false,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                 PropertyPaneToggle("showIndicators", {
                  label: strings.showIndicatorsFieldLabel,
                  checked: true,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                 PropertyPaneToggle("infiniteLoop", {
                  label: strings.infiniteLoopFieldLabel,
                  checked: true,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                PropertyPaneToggle("swipeable", {
                  label: strings.swipeableFieldLabel,
                  checked: true,
                  onText: "Toggle is On", offText:"Toggle is Off"
                }),
                PropertyPaneSlider("interval", {
                  label: strings.intervalFieldLabel,
                  min: 1000,
                  max: 10000,
                  step: 500,
                  value: 3000,
                 }),
                 PropertyPaneSlider("transitionTime", {
                  label: strings.transitionTimeFieldLabel,
                  min: 500,
                  max: 3000,
                  step: 500,
                  value: 500,
                 }),
              ]
            }
          ]
        }
      ]
    };
  }
}
