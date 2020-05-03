import * as React from 'react';
import styles from './ExcelChartsSlider.module.scss';
import { IExcelChartsSliderProps } from './IExcelChartsSliderProps';
import { IExcelChartsSliderState } from './IExcelChartsSliderState';
import { IExcelServices } from '../Services/IExcelServices';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ISheetKeywords } from '../Services/ISheetKeywords';
//const logo: any = require('../assets/charticon.jpg');


export default class ExcelChartsSlider extends React.Component<IExcelChartsSliderProps, IExcelChartsSliderState> {

  constructor(props: IExcelChartsSliderProps, state:IExcelChartsSliderState)
  {
      super(props);
      this.state = {items: [],
      Kitems:[],
      loading: false,
      selectedKeyword:"",
      selectedChart:"",
      active:null,
      activeChart:null,
      spHttpClient: this.props.spHttpClient,
      title: this.props.title,
      wbId: this.props.wbId,
};

 }

  public componentDidMount(): void {
      this._getSheets().then((result: Array<ISheetKeywords>) => {
      this.setState({Kitems: result, loading: false});
    });
  }

  public render(): React.ReactElement<any> {
          return (

      <div className={ styles.excelChartsSlider }>
        <div className={ styles.container }>
            <span className={styles.title}><b>{this.props.title}</b></span>
            <br></br>

            {this.state.Kitems.length && this.state.Kitems.map((listItemT, index) => {

              return (
                <div className={styles.keywordsDiv}>
                  <span key={index} id={index.toString()}
                    style={{ background: this._myColor(index) }}
                    className={styles.KW} onClick={(event) => this._filterbyKeyword(event, listItemT.Keywords, index)}>{listItemT.Keywords}</span>
                </div>
              );
            })}
            <br></br>
            <hr className={styles.Charthr}></hr>
            <div id="ImageDiv"><img id="ChartSrc" src="" alt=""/></div>
            <hr className={styles.Charthr}></hr>
            <div id="ChartSpans">
            {this.state.items.length && this.state.items.map((listItem, index) => {
              return (

                    <span key={index} id={index.toString()}
                    style={{ background: this._myChartColor(index) }}
                    className={styles.KW} onClick={(event) => this._filterbyCharts(event, listItem.Title, index)}>{listItem.Title}</span>

                    );
                })}
            </div>
              <div>
           </div>
        </div>
      </div>
    );
  }


  public _myColor(index) {
    if (this.state.active === index) {
      return "#7e159e";
    }
    return "#0078d4";
  }

  public _myChartColor(index) {
    if (this.state.activeChart === index) {
      return "#7e159e";
    }
    return "#0078d4";
  }

  public _filterbyKeyword(event, keywordName, index) {
    (document.getElementById('ChartSrc') as HTMLImageElement).src = "";

    var nodes = document.getElementById('ChartSpans').getElementsByTagName("span");
    for (var i = 0; i < nodes.length; i++) {
      nodes[i].style.background = '#0078d4';
    }
    this.setState({ selectedKeyword: keywordName });
    if (this.state.active === index) {
      this.setState({ active: null });
    } else {
      this.setState({ active: index });
    }
    this._getCharts(keywordName).then((result: Array<IExcelServices>) => {
      this.setState({ items: result,loading: false });
    });

  }

  public _filterbyCharts(event, chartName, index) {
    this.setState({ loading: true });
    this.setState({ selectedChart: chartName });
    if (this.state.activeChart === index) {
      this.setState({ activeChart: null });
    } else {
      this.setState({ activeChart: index });
    }
    this.props.graphClient.api("/me/drive/items/" + this.props.wbId + "/workbook/worksheets('" + this.state.selectedKeyword + "')/charts('" + chartName + "')/image")
    .get((error,response: any) => {
      if (error) {
        console.log(error);
        return;
      }
      if (response && response.value)
      {
      (document.getElementById('ChartSrc') as HTMLImageElement).src = "data:image/png;base64," + response.value;
      }
      else
      {
        console.log('no message found');
      }
    });
  }



  public _getCharts(skw:string,options?: any,): Promise<IExcelServices[]> {
    return new Promise<IExcelServices[]>((resolve: any) => {

        var wbid = this.props.wbId;
        if (skw !== "")
        {
          this.props.graphClient.api("/me/drive/items/" + wbid + "/workbook/worksheets/" + skw + "/charts")
              .get((error,response: any,charts:MicrosoftGraph.WorkbookChart[], rawResponse?: any) => {
                if (error) {
                  console.log(error);
                  return;
                }
                if (response && response.value && response.value.length > 0)
                {
                  this.setState({ loading: true });
                   charts = response.value;
                  let itemsCh: IExcelServices[] = [];
                    for (let index = 0; index < charts.length; index++)
                    {
                      itemsCh.push({Title:charts[index].name});

                    }
                    this.setState({items: itemsCh});
                    resolve(itemsCh);
                }

                else {
                  console.log('no message found');
                }
              });

        }
    });
  }

  public _getSheets(options?: any): Promise<ISheetKeywords[]> {
    return new Promise<ISheetKeywords[]>((resolve: any) => {
      var wbid = this.props.wbId;
      if (wbid !== "")
      {

          this.props.graphClient.api("/me/drive/items/" + wbid + "/workbook/worksheets")
          .get((error,response: any, rawResponse?: any) => {
              if (error) {
                console.error(error);
                return;
              }
              if (response && response.value && response.value.length > 0) {
                this.setState({ loading: true });
                let sheets:MicrosoftGraph.WorkbookWorksheet[] = response.value;
                let itemsKeywords: ISheetKeywords[] = [];

                for (let index = 0; index < sheets.length; index++) {
                  // Populate array with the excel wb sheet names
                  itemsKeywords.push({ Keywords: sheets[index].name});
                }
                this.setState({ Kitems: itemsKeywords });
                resolve(itemsKeywords);
              }
              else {
                //01BXEHNNVQ6UMAL677GRALHFD6FQMZTYKY
                console.log('no message found');
              }

        });
      }
    });
  }
}
