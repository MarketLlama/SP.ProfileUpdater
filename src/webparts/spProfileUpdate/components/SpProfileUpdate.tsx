import * as React from 'react';
import { sp } from "@pnp/sp";
import { Fabric } from 'office-ui-fabric-react/lib/Fabric'
import styles from './SpProfileUpdate.module.scss';
import { ISpProfileUpdateProps } from './ISpProfileUpdateProps';
import { ISpProfileUpdateState } from './ISpProfileUpdateState'
import { Collapse } from 'react-collapse';
import { IconButton, IButtonProps, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { 
  TaxonomyPicker, 
  IPickerTerms , 
  IPickerTerm
} from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

import {
  taxonomy,
  ITermStore,
  ITerms,
  ILabelMatchInfo,
  ITerm,
  ITermData,
  ITermStoreData,
  ITermSet,
  ITermSetData,
} from "@pnp/sp-taxonomy";

export default class SpProfileUpdate extends React.Component<ISpProfileUpdateProps, ISpProfileUpdateState> {

  constructor(props, state) {
    super(props);

    this.state = {
      time: new Date(),
      firstName : '',
      imageUrl: '',
      accountName : '',
      title : '',
      termList : [],
      open : false,
      defaultLocationTerms : [],
      defaultDepartmentTerms : [],
      defaultLanguageTerms : []
    };
    this.onInit.bind(this);
    this.onInit();
    this.updateProfile.bind(this);
  }

  public onInit() :void{
    sp.profiles.myProperties.get().then(user =>{
      console.log(user);

      let imageUrl = (user.PictureUrl)? user.PictureUrl.replace("MThumb","LThumb") : 
        this.props.context.pageContext.site.absoluteUrl + "/_layouts/15/images/PersonPlaceholder.200x150x32.png";
      let accountName = user.AccountName;
      let title = user.Title;
      //get firstName
      let firstName :string = user.UserProfileProperties.find(item =>{
        return item.Key == "FirstName";
      }).Value;

      let location :string = user.UserProfileProperties.find(item =>{
          return item.Key == "SPS-Location";
      }).Value;
      if(location){
        this.getInitialValues(location).then(terms =>{
          this.setState({
            defaultLocationTerms : terms
          });
        });
      }

      let department :string = user.UserProfileProperties.find(item =>{
          return item.Key == "SPS-Department";
      }).Value;
      if(department){
        this.getInitialValues(department).then(terms =>{
          this.setState({
            defaultDepartmentTerms : terms
          });
        });
      }
            
      let language :string = user.UserProfileProperties.find(item =>{
        return item.Key == "SPS-MUILanguages";
      }).Value;
      if(language){
        this.getInitialValues(language).then(terms =>{
          this.setState({
            defaultLanguageTerms : terms
          });
        });
      }

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
    if(h >= 12){
      greeting = 'Afternoon, ';
    }
    if(h > 18){
      greeting = 'Evening, ';
    }
    return (
      <Fabric>
      <div className={styles.spProfileUpdate}>
        <div className={[styles.login_box, styles.row].join(' ')}>
            <div className={[styles["col-md-12"],styles["col-sm-12"]].join(' ')} style={{textAlign:'center'}}>
                  <div className={styles.line}><h3 style={{padding: '10px', marginTop:0,fontWeight:600}}>{(h % 12) == 0? '12' : h % 12}:{(m < 10 ? '0' + m : m)}:{(s < 10 ? '0' + s : s)} {h < 12 ? 'AM' : 'PM'}</h3></div>
                  <div className={styles.outter}><img src={this.state.imageUrl} className={styles["image-circle"]}/></div>   
                  <h1>{greeting}{this.state.firstName}</h1>
                  <span>{this.state.title}</span>
                  <br/>
                  <IconButton iconProps={(this.state.open)? { iconName: 'DoubleChevronUp' } : { iconName: 'DoubleChevronDown' }} 
                  onClick={() => this.setState({ open: !this.state.open })}/>
            </div>    
        </div>
          <Collapse isOpened={this.state.open}>
            <div style={{padding:'20px'}} className={styles.row}>
                <TaxonomyPicker
                    allowMultipleSelections={false}
                    termsetNameOrID="Wizdom_Languages"
                    panelTitle="Select Language"
                    label="Language"
                    context={this.props.context}
                    onChange={this.onTaxChangeLanguage}
                    isTermSetSelectable={false}
                    initialValues={this.state.defaultLanguageTerms}
                  />
                <TaxonomyPicker
                    allowMultipleSelections={false}
                    termsetNameOrID="GeographyHierarchy"
                    panelTitle="Select Location"
                    label="Location"
                    context={this.props.context}
                    onChange={this.onTaxChangeLocation}
                    isTermSetSelectable={false}
                    initialValues={this.state.defaultLocationTerms}
                  />
                <TaxonomyPicker
                    allowMultipleSelections={false}
                    termsetNameOrID="ApplicableFunction"
                    panelTitle="Select Department"
                    label="Department"
                    context={this.props.context}
                    onChange={this.onTaxChangeDepartment}
                    isTermSetSelectable={false}
                    initialValues={this.state.defaultDepartmentTerms}
                  />
                <PrimaryButton disabled={this.state.termList.length == 0 } 
                    iconProps={{ iconName: 'AddFriend' }}
                    onClick={this.updateProfile.bind(this)} 
                    style={{float:'right', margin:'20px'}}>Update</PrimaryButton>
                <br/>
              </div>
          </Collapse>
      </div>
      </Fabric>);
     
  }

  public componentDidMount() {
    setInterval(this.updateTime.bind(this), 1000);
	}
	
  private updateTime() {	
		this.setState({
			time: new Date()
		});
  }
  
  //Best way i could do it....
  private onTaxChangeLanguage = (terms : IPickerTerms) => {
    this.onTaxChange(terms, "Wizdom_Languages");
  } 

  private onTaxChangeLocation = (terms : IPickerTerms) => {
    this.onTaxChange(terms, "GeographyHierarchy");
  } 

  private onTaxChangeDepartment = (terms : IPickerTerms) => {
    this.onTaxChange(terms, "ApplicableFunction");
  } 

  private onTaxChange = (terms : IPickerTerms, termSetName : string) => {
    let newTermList : IPickerTerms = [];
    if(terms.length !== 0){
      newTermList.push(terms[0]);
    } else {
      if(this.state.termList.length !== 0){
        newTermList = this.state.termList.filter(term =>{
          return term.termSetName == termSetName
        });
      }
    }
    this.setState({
      termList : newTermList
    });
  }

  private updateProfile = () => {
    let batch = sp.createBatch();

    let department = this.state.termList.filter(term => { return term.termSetName == "ApplicableFunction"})[0];
    let location = this.state.termList.filter(term => { return term.termSetName == "GeographyHierarchy"})[0];
    let language = this.state.termList.filter(term => { return term.termSetName == "Wizdom_Languages"})[0];

    if(department){
      sp.profiles.inBatch(batch).setSingleValueProfileProperty(this.state.accountName,"SPS-Department",department.name);
    }
    if(location){
      sp.profiles.inBatch(batch).setSingleValueProfileProperty(this.state.accountName,"SPS-Location",location.name);
    }
    if(language){
      sp.profiles.inBatch(batch).setSingleValueProfileProperty(this.state.accountName,"SPS-MUILanguages",language.name);
    }

    batch.execute().then(() => {
      alert("All done!");
    }, f =>{
      console.log(f);
    });
  }

  private async getInitialValues(userProfileProp : string){
    let userTerms : IPickerTerms = [];
    
    let store: (ITermStoreData & ITermStore)[] = await taxonomy.termStores.get();

    let labelMatchInfo: ILabelMatchInfo = {
      TermLabel: userProfileProp,
      TrimUnavailable: true,
    };
  
    let terms: (ITermData & ITerm)[] = await store[0].getTerms(labelMatchInfo).get();

    let regExp = /\(([^)]+)\)/;
    for(let term of terms){
        let termSet :(ITermSetData & ITermSet) =  await this.getTermSet(term);
        userTerms.push({
          key : regExp.exec(term.Id)[1],
          name : term.Name,
          path : term.PathOfTerm,
          termSet : regExp.exec(termSet.Id)[1]
      });
    }
    return userTerms;
  }

  private async getTermSet(term : ITermData & ITerm){
    return term.termSet.get();
  }
}
