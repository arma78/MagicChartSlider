import * as React from 'react';
import styles from './ExcelChartsSlider.module.scss';
import { IExcelChartsSliderProps } from './IExcelChartsSliderProps';
import { IExcelChartsSliderState } from './IExcelChartsSliderState';
import { IExcelServices } from '../Services/IExcelServices';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ISheetKeywords } from '../Services/ISheetKeywords';
import { Shimmer} from 'office-ui-fabric-react/lib/Shimmer';

const chartIcon2: any = require('./assets/charticon2.png');
const chartIcon1: any = require('./assets/charticon1.png');
const chartIcon: any = require('./assets/charticon.png');
const spinner: any = require('./assets/spinner.gif');
export default class ExcelChartsSlider extends React.Component<IExcelChartsSliderProps, IExcelChartsSliderState> {


  constructor(props: IExcelChartsSliderProps, state: IExcelChartsSliderState) {
    super(props);
    this._filterbyKeyword = this._filterbyKeyword.bind(this);
    this._filterbyCharts = this._filterbyCharts.bind(this);
    this.state = {
      items: [],
      Kitems: [],
      loading: true,
      loadingChart: true,
      selectedKeyword: "",
      selectedChart: "",
      active: null,
      activeChart: null,
      spHttpClient: this.props.spHttpClient,
      title: this.props.title,
      wbId: this.props.wbId,
      chart: "",
      chartspans:"none",
      imagediv:"block",
    };
  }

  public componentDidMount(): void {

    if (this.props.wbId !== "" || this.props.wbId !== null)
    {

    this._getSheets().then((result: Array<ISheetKeywords>) => {
      this.setState({ Kitems: result,chart:chartIcon2});
    });

    if (this.props.wbId == "" || this.props.wbId == null)
    {

      this.setState({chart: chartIcon1});
    }
  }
  }
  public componentDidUpdate(prevProps, prevState) {



    if(this.state.loading !== prevState.loading )
    {
      this._shimmer();
    }
    if(this.state.loadingChart !== prevState.loadingChart )
    {
      this._shimmerCharts();
    }
    if (this.props.wbId !== prevProps.wbId) {
          this.setState({Kitems:[],items:[],active:null,activeChart:null, chart:spinner});
          this._getSheets().then((result: Array<ISheetKeywords>) => {
          this.setState({ Kitems: result,chart:chartIcon2});
        });
     }
  }

  public render(): React.ReactElement<IExcelChartsSliderProps> {

     return (

      <div className={styles.excelChartsSlider}>
        <div className={styles.container}>
          <span className={styles.title}><b>{this.props.title}</b></span>
          <br></br>

          {this.state.Kitems.length && this.state.Kitems.map((listItemT, index) => {

            return (
              <div className={styles.keywordsDiv} id="keywordsDiv">
                <span key={index} id={index.toString()}
                  style={{ background: this._myColor(index),pointerEvents: this._DisableClick(index) }}
                  className={styles.KW} onClick={(event) => this._filterbyKeyword(event, listItemT.Keywords, index)}>{listItemT.Keywords}</span>
              </div>
            );
          })}
          <br></br>
          <hr className={styles.Charthr}></hr>
          <Shimmer  isDataLoaded={this.state.loadingChart}>
          </Shimmer>
          <div  style={{display:this.state.imagediv }} id="ImageDiv">
            <img id="ChartSrc" src={this.state.chart} alt="" />
          </div>
          <hr className={styles.Charthr}></hr>
          <Shimmer isDataLoaded={this.state.loading}>
          </Shimmer>
          <div id="ChartSpans"  style={{display:this.state.chartspans}}>
            {this.state.items.length && this.state.items.map((listItem, index) => {
              return (
                <span key={index} id={index.toString()}
                  style={{ background: this._myChartColor(index), pointerEvents: this._DisableClickCharts(index) }}
                  className={styles.KW} onClick={(event) => this._filterbyCharts(event, listItem.Title, index)}>{listItem.EnumTitle}</span>
              );
            })}
          </div>
          <div>
          </div>
        </div>
      </div>
    );
  }

  private _shimmer() {


    if (this.state.loading == false) {
      return  this.setState({chartspans:"none"});
    }
    else {
     return  this.setState({chartspans:"block"});
    }
  }
  private _shimmerCharts() {

    if (this.state.loadingChart == false) {
      return this.setState({imagediv:"none"});
    }
    else {
      return this.setState({imagediv:"block"});
    }
  }


  private _myColor(index) {

    if (this.state.active === index) {
      return "#7e159e";
    }
    return "#0078d4";
  }
  private _DisableClick(index)
  {
    if (this.state.active === index) {
      return "none";
    }
    return "auto";

  }
  private _DisableClickCharts(index)
  {
    if (this.state.activeChart === index) {
      return "none";
    }
    return "auto";

  }

  private _myChartColor(index) {
    if (this.state.activeChart === index) {
      return "#7e159e";
    }
    return "#0078d4";
  }

  private _filterbyKeyword(event, keywordName, index) {
    this.setState({loading:false,
      activeChart:null,
       selectedKeyword: keywordName,
       chart:"" });
    if (this.state.active === index) {
      this.setState({ active: null });
    } else {
      this.setState({ active: index });
    }
    this._getCharts(keywordName).then((result: Array<IExcelServices>) => {
      this.setState({ items: result,loading:true});
    });
  }

  private _filterbyCharts(event, chartName, index) {
    this.setState({ selectedChart: chartName, loadingChart: false });
    if (this.state.activeChart === index) {
      this.setState({ activeChart: null });
    } else {
      this.setState({ activeChart: index });
    }
    this.props.graphClient.api("/me/drive/items/" + this.props.wbId + "/workbook/worksheets('" + this.state.selectedKeyword + "')/charts('" + chartName + "')/image")
      .get((error, response: any) => {
        if (error) {
          console.log(error);
          return;
        }
        if (response && response.value) {
          this.setState({chart:"data:image/png;base64," + response.value });
        }
        else {
          console.log('Chart not found in this WorkSheet');
        }
        this.setState({loadingChart:true });
      });
  }



  private _getCharts(skw: string, options?: any,): Promise<IExcelServices[]> {
    return new Promise<IExcelServices[]>((resolve: any) => {
      var wbid = this.props.wbId;
      if (skw !== "") {
        this.props.graphClient.api("/me/drive/items/" + wbid + "/workbook/worksheets/" + skw + "/charts")
          .get((error, response: any, charts: MicrosoftGraph.WorkbookChart[], rawResponse?: any) => {
            if (error) {
              console.log(error);
              return;
            }
            if (response && response.value && response.value.length > 0) {
              charts = response.value;
              let itemsCh: IExcelServices[] = [];
              for (let index = 0; index < charts.length; index++) {
                itemsCh.push({Title:charts[index].name ,EnumTitle: "Chart - " + (index + 1).toString()});
              }
              this.setState({ items: itemsCh });
              resolve(itemsCh);
            }
            else {
              this.setState({ items:[], chart:chartIcon });
              console.log('no charts found');
            }
          });
      }
    });
  }

  private _getSheets(options?: any): Promise<ISheetKeywords[]> {
    return new Promise<ISheetKeywords[]>((resolve: any) => {
      var wbid = this.props.wbId;
      if (wbid !== "") {

        this.props.graphClient.api("/me/drive/items/" + wbid + "/workbook/worksheets")
          .get((error, response: any) => {
            if (error) {
              console.error(error);
              return;
            }
            if (response && response.value && response.value.length > 0) {
              let sheets: MicrosoftGraph.WorkbookWorksheet[] = response.value;
              let itemsKeywords: ISheetKeywords[] = [];

              for (let index = 0; index < sheets.length; index++) {
                // Populate array with the excel wb sheet names
                itemsKeywords.push({ Keywords: sheets[index].name });
              }
              this.setState({ Kitems: itemsKeywords});
              resolve(itemsKeywords);
            }
            else {
              console.log('No Sheets in WorkBook');
            }

          });
      }
    });
  }
}
