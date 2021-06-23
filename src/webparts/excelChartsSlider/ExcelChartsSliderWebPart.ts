
import { IExcelWBList } from './Services/IExcelWBList';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {SPHttpClient} from "@microsoft/sp-http";
import { Version} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
}
from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
}
from '@microsoft/sp-property-pane';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
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
export interface IPropertyPaneDropdownOption
{
  key:string;
  text:string;
}


export default class ExcelChartsSliderWebPart extends BaseClientSideWebPart<IExcelChartsSliderWebPartProps>
{
  private listDropDownOptions: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;







  private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
    let items: IExcelWBList[] = [];


    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject) => {
      let options: IPropertyPaneDropdownOption[] = [];
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client.api("/me/drive/search(q='.xlsx')?select=id,name")
            .orderby("name")
            // tslint:disable-next-line:no-shadowed-variable
            .get((error, response: any, rawResponse?: any) => {
              if (response && response.value && response.value.length > 0) {
                let drive: MicrosoftGraph.DriveItem[];

                drive = response.value;
                for (let index = 0; index < drive.length; index++) {
                  items.push({ Id: drive[index].id, name: drive[index].name });
                }
                items.map((list: IExcelWBList) => {
                  options.push({ key: list.Id, text: list.name });
                });
              } else {
                console.log(error);
                return null;
              }
              resolve(options);
            });
        });
    });

  }






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






  protected onPropertyPaneConfigurationStart(): void
  {



    //Bind DropDown List in Peropert pane
    this.listsDropdownDisabled = !this.listDropDownOptions;
    if (this.listDropDownOptions)
    {
      return;
    }
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'listDropDownOptions');


       this.fetchOptions()
      .then((listsResp: IPropertyPaneDropdownOption[]):void => {
        this.listDropDownOptions = listsResp;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
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
                PropertyPaneDropdown('wbId',{
                  label: "Select Excel WorkBook",
                  options:this.listDropDownOptions,
                  disabled: this.listsDropdownDisabled
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
