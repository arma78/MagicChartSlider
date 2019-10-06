import {IExcelServices } from "../Services/IExcelServices";

import {SPHttpClient} from "@microsoft/sp-http";
export interface IExcelChartsSliderState {
  items: IExcelServices[];
  listName: string;
  siteurl: string;
  spHttpClient: SPHttpClient;
  title: string;
  wbName:string;
}
