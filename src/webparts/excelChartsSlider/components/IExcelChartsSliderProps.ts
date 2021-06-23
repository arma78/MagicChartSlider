import {SPHttpClient, MSGraphClient } from '@microsoft/sp-http';
export interface IExcelChartsSliderProps {

  wbId: string;
  description: string;
  title: string;
  spHttpClient: SPHttpClient;
  graphClient: MSGraphClient;
}
