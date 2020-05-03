import * as React from 'react';
import * as ReactDom from 'react-dom';
import {SPHttpClient} from "@microsoft/sp-http";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
}
from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
}
from '@microsoft/sp-property-pane';
import { MSGraphClient } from '@microsoft/sp-http';
import * as strings from 'ExcelChartsSliderWebPartStrings';
import ExcelChartsSlider from './components/ExcelChartsSlider';
import { IExcelChartsSliderProps } from './components/IExcelChartsSliderProps';

export interface IExcelChartsSliderWebPartProps {

  description: string;
  wbId: string;
  spHttpClient: SPHttpClient;
  title: string;

}

export default class ExcelChartsSliderWebPart extends BaseClientSideWebPart<IExcelChartsSliderWebPartProps> {

  public render(): void {
    this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
    const element: React.ReactElement<IExcelChartsSliderProps > = React.createElement(
      ExcelChartsSlider,
      {
        title:this.properties.title,
        description: this.properties.description,
        wbId: this.properties.wbId,
        spHttpClient: this.properties.spHttpClient,
        graphClient: client,
      }
    );

    ReactDom.render(element, this.domElement);
  });
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
                PropertyPaneTextField('wbId', {
                  label: strings.wbIdFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
