import {SPHttpClient, } from "@microsoft/sp-http";
import { MSGraphClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IExcelChartsSliderProps {

  wbId: string;
  description: string;
  title: string;
  spHttpClient: SPHttpClient;
  graphClient: MSGraphClient;
  //context: WebPartContext;
}
