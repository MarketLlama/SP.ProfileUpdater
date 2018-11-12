import * as React from 'react';
import styles from './SpProfileUpdate.module.scss';
import { ISpProfileUpdateProps } from './ISpProfileUpdateProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Button from 'react-bootstrap/lib/Button';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

export default class SpProfileUpdate extends React.Component<ISpProfileUpdateProps, {}> {
  constructor(props) {
    super(props);
  }

  public componentWillMount() : void {

  }

  public render(): React.ReactElement<ISpProfileUpdateProps> {
    return (
      <div className={ styles.spProfileUpdate }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
              <Button>Hello World</Button>
              <TaxonomyPicker
                  allowMultipleSelections={false}
                  termsetNameOrID="ApplicableFunction"
                  panelTitle="Select Function"
                  label="Taxonomy Picker"
                  context={this.props.context}
                  onChange={this.onTaxPickerChange}
                  isTermSetSelectable={false}
                />
            </div>
          </div>
        </div>
      </div>
    );
    /*
        return (
      <div className={[styles.spProfileUpdater, styles.row].join(' ')}>
        <div className={[styles.row, styles.login_box].join(' ')}>
            <div className={[styles["col-md-12"],styles["col-sm-12"]].join(' ')} style={{textAlign:'center'}}>
                  <div className={styles.line}><h3>12 : 30 AM</h3></div>
                  <div className={styles.outter}><img src="http://lorempixel.com/output/people-q-c-100-100-1.jpg" className={styles["image-circle"]}/></div>   
                  <h1>Hi Guest</h1>
                  <span>INDIAN</span>
            </div>
              <div className={[styles["col-md-6"], styles["col-sm-6"], styles.follow ,styles.line].join(' ')} style={{textAlign:'center'}}>
                  <h3>
                      125651 <br/> <span>FOLLOWERS</span>
                  </h3>
              </div>
              <div className={[styles["col-md-6"], styles["col-sm-6"], styles.follow, styles.line].join(' ')} style={{textAlign:'center'}}>
                  <h3>
                      125651 <br/> <span>FOLLOWERS</span>
                  </h3>
              </div>
              <div className={[styles["col-md-12"], styles["col-sm-12"], styles.login_control].join(' ')}>
                <TermPicker/>
              </div>       
          </div>
      </div>);
     */
  }

  public componentDidMount() :void{

  }


  private onTaxPickerChange() : void{
    console.log("Hello");
  }


}
