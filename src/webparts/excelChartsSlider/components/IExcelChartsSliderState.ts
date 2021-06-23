import {IExcelServices } from "../Services/IExcelServices";
import {ISheetKeywords } from "../Services/ISheetKeywords";
import {SPHttpClient} from "@microsoft/sp-http";
export interface IExcelChartsSliderState {
  items: IExcelServices[];
  Kitems: ISheetKeywords[];
  spHttpClient: SPHttpClient;
  title: string;
  wbId: string;
  loading: boolean;
  loadingChart: boolean;
  selectedKeyword:string;
  selectedChart:string;
  active: string;
  activeChart: string;
  chart:string;
  chartspans:string;
  imagediv:string;
}
