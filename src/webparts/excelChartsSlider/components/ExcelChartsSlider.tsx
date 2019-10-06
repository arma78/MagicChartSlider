import * as React from 'react';
import styles from './ExcelChartsSlider.module.scss';
import { IExcelChartsSliderProps } from './IExcelChartsSliderProps';
import { IExcelChartsSliderState } from './IExcelChartsSliderState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Carousel } from '../../../../node_modules/react-responsive-carousel';
import "react-responsive-carousel/lib/styles/carousel.min.css";
import { IExcelServices } from '../Services/IExcelServices';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class ExcelChartsSlider extends React.Component<IExcelChartsSliderProps, IExcelChartsSliderState> {
  constructor(props: IExcelChartsSliderProps, state:IExcelChartsSliderState)
  {
      super(props);
      this.state = {items: [],
      siteurl: this.props.siteurl,
      spHttpClient: this.props.spHttpClient,
      title: this.props.title,
      listName: this.props.listName,
      wbName:this.props.wbName};

 }

  public componentDidMount(): void {
    this._getCharts().then((result: Array<IExcelServices>) => {
      this.setState({items: result});
    });
  }
  public componentDidUpdate(prevState, prevProps) {

    if (this.props.listName !== prevProps.listName ||
        this.props.wbName !== prevProps.wbName) {
          this.setState({  title: this.props.title,
            listName: this.props.listName,
            wbName:this.props.wbName});

          this._getCharts().then((result: Array<IExcelServices>) => {
            this.setState({items: result });
          });
    }
  }

  public render(): React.ReactElement<IExcelChartsSliderProps> {
    return (
      <div className={ styles.excelChartsSlider }>
        <div className={ styles.container }>
              <span className={styles.title}><b>{this.props.title}</b></span>
              <Carousel
                showThumbs={this.props.showThumbs}
                autoPlay={this.props.autoPlay}
                infiniteLoop={this.props.infiniteLoop}
                interval={this.props.interval}
                showArrows={this.props.showArrows}
                showStatus={this.props.showStatus}
                swipeable={this.props.swipeable}
                stopOnHover={this.props.stopOnHover}
                showIndicators={this.props.showIndicators}
                transitionTime={this.props.transitionTime}>
                {this.state.items.length && this.state.items.map((listItem, index) => {
                  return (
                    <div>
                      <img key={index} src={this.state.siteurl + "/_vti_bin/ExcelRest.aspx/" + this.state.listName + "/" + this.state.wbName + "/Model/Charts('" + listItem.Title + "')?$format=image"} />
                    </div>);
                })}
              </Carousel>
        </div>
      </div>
    );
  }

  public _getCharts(options?: any): Promise<IExcelServices[]> {
    return new Promise<IExcelServices[]>((resolve: any) => {
      fetch(this.state.siteurl + "/_vti_bin/ExcelRest.aspx/" + this.state.listName + "/" + this.state.wbName + "/Model/Charts?$format=json")
        .then(res => res.json())
        .then((data) => {
          let itemsCh: IExcelServices[] = [];
          for (var i = 0; i < data.items.length; i++) {
            itemsCh.push({Title : data.items[i].name});
          }
          this.setState({ items: itemsCh });
          resolve(itemsCh);
        });
    });
  }

}
