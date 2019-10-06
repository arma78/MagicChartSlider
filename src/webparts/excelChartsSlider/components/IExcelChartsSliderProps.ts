import {SPHttpClient} from "@microsoft/sp-http";

export interface IExcelChartsSliderProps {
  listName: string;
  description: string;
  siteurl: string;
  wbName: string;
  title: string;
  spHttpClient: SPHttpClient;
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
