import * as React from 'react';
import styles from './SpProfileUpdate.module.scss';
import { ISpProfileUpdateProps } from './ISpProfileUpdateProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Button from 'react-bootstrap/lib/Button';
import Well from 'react-bootstrap/lib/Well';
import Panel from 'react-bootstrap/lib/Panel';
import PanelGroup from 'react-bootstrap/lib/PanelGroup';
import accordian from 'react-bootstrap/lib/Accordion';
import collapsible from 'react-bootstrap/lib/Collapse';
import { TaxonomyPicker, IPickerTerms , IPickerTerm} from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { sp } from "@pnp/sp";

		
export interface ISpProfileUpdateState{
  time : Date;
  firstName : string;
  imageUrl : string;
  title : string;
  accountName : string;
}

export default class SpProfileUpdate extends React.Component<ISpProfileUpdateProps, ISpProfileUpdateState> {
  private _terms: any[];
  public get terms(): any[] {
    return this._terms;
  }
  public set terms(value: any[]) {
    this._terms = value;
  }
  constructor(props, state:ISpProfileUpdateState) {
    super(props);

    this.state = {
      time: new Date(),
      firstName : '',
      imageUrl: '',
      accountName : '',
      title : ''
    };
    this.onInit.bind(this);
    this.onInit();
    this.updateProfile.bind(this);
    this.pushTerms.bind(this);
    this.popTerms.bind(this);
  }

  public onInit() :void{
    sp.profiles.myProperties.get().then(user =>{
      console.log(user);
      let imageUrl = (user.PictureUrl)? user.PictureUrl : 
        this.props.context.pageContext.site.absoluteUrl + "/_layouts/15/images/PersonPlaceholder.200x150x32.png";
      let accountName = user.AccountName;
      let title = user.Title;
      //get firstName
      let firstName :string = user.UserProfileProperties.find(item =>{
        return item.Key == "FirstName";
      }).Value;
      this.setState({
        firstName : firstName,
        title : title,
        accountName : accountName,
        imageUrl : imageUrl
      });
    }, error =>{
      console.log(error);
    });
  }

  public render(): React.ReactElement<ISpProfileUpdateProps> {
		const h = this.state.time.getHours();
		const m = this.state.time.getMinutes();
    const s = this.state.time.getSeconds();
    let greeting = '';
    if(h < 12){
      greeting = 'Morning, ';
    } 
    if(h > 12){
      greeting = 'Afternoon, ';
    }
    if(h > 18){
      greeting = 'Evening, ';
    }
    return (
      <div className={styles.spProfileUpdate}>
        <div className={[styles.row, styles.login_box].join(' ')}>
            <div className={[styles["col-md-12"],styles["col-sm-12"]].join(' ')} style={{textAlign:'center'}}>
                  <div className={styles.line}><h3>{h % 12}:{(m < 10 ? '0' + m : m)}:{(s < 10 ? '0' + s : s)} {h < 12 ? 'AM' : 'PM'}</h3></div>
                  <div className={styles.outter}><img src={this.state.imageUrl} className={styles["image-circle"]}/></div>   
                  <h1>{greeting}{this.state.firstName}</h1>
                  <span>{this.state.title}</span>
            </div>    
          </div>
          <div>
            <PanelGroup accordion id="accordion-example">
              <Panel eventKey="1">
                <Panel.Heading>
                  <Panel.Title toggle>Collapsible Group Item #1</Panel.Title>
                </Panel.Heading>
                <Panel.Body collapsible>
                  <Well>
                  <TaxonomyPicker
                      allowMultipleSelections={false}
                      termsetNameOrID="Wizdom_Languages"
                      panelTitle="Select Language"
                      label="Language"
                      context={this.props.context}
                      onChange={this.onLanguageChange}
                      isTermSetSelectable={false}
                    />
                  </Well>
                </Panel.Body>
              </Panel>
              <Panel eventKey="2">
                <Panel.Heading>
                  <Panel.Title toggle>Collapsible Group Item #2</Panel.Title>
                </Panel.Heading>
                <Panel.Body collapsible>
                  <Well>
                  <TaxonomyPicker
                      allowMultipleSelections={false}
                      termsetNameOrID="GeographyHierarchy"
                      panelTitle="Select Location"
                      label="Location"
                      context={this.props.context}
                      onChange={this.onLocationChange}
                      isTermSetSelectable={false}
                    />
                  </Well>
                </Panel.Body>
              </Panel>
              <Panel eventKey="3">
                <Panel.Heading>
                  <Panel.Title toggle>Collapsible Group Item #3</Panel.Title>
                </Panel.Heading>
                <Panel.Body collapsible>
                  <Well>
                      <TaxonomyPicker
                      allowMultipleSelections={false}
                      termsetNameOrID="ApplicableFunction"
                      panelTitle="Select Department"
                      label="Department"
                      context={this.props.context}
                      onChange={this.onDepartmentChange}
                      isTermSetSelectable={false}
                    />
                  </Well>
                </Panel.Body>
              </Panel>
            </PanelGroup>
            <Button style={{float:'right'}}>Update</Button>
            <br/>
          </div>
      </div>);
     
  }

  public componentDidMount() {
		setInterval(this.updateTime.bind(this), 1000);
	}
	
  private updateTime() {	
		this.setState({
			time: new Date()
		});
	}

  private onLanguageChange(terms : IPickerTerms) : void{
    console.log(terms);
    /*if(item){
      this.terms(item[0]);
    }else {

    }*/
  }
  private onLocationChange(terms : IPickerTerms) : void{
    console.log(terms);
    /*if(item){
      this.terms(item[0]);
    }else {

    }*/
  }
  private onDepartmentChange(terms : IPickerTerms) : void{
    console.log(terms);
    if(terms){
      this.pushTerms(terms[0]);
    }else {
      this.popTerms("ApplicableFunction");
    }
  }
  private pushTerms(term : IPickerTerm):void{
    this._terms.push(term);
  }

  private popTerms(termSetName : string):void{
    let array = this._terms.filter(term =>{
      return term.termSetName == termSetName;
    });

    console.log(array);
  }
  private updateProfile():void{
    sp.profiles.setSingleValueProfileProperty(this.state.accountName,"","").then(item =>{

    },error =>{

    });

  }


}
