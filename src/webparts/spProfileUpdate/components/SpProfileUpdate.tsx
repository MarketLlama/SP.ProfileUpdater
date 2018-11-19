import * as React from 'react';
import { sp } from "@pnp/sp";
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import styles from './SpProfileUpdate.module.scss';
import { ISpProfileUpdateProps } from './ISpProfileUpdateProps';
import { ISpProfileUpdateState } from './ISpProfileUpdateState';
import { Collapse } from 'react-collapse';
import {SpProfileReminder} from './SpProfileReminder';
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
  private _targetButton = null;
  public setTarget: (element: any) => void;

  constructor(props, state) {
    super(props);
    this.state = {
      time: new Date(),
      firstName : '',
      imageUrl: '',
      accountName : '',
      title : '',
      termList : [],
      showCheckMark : false,
      open : false,
      defaultLocationTerms : [],
      defaultDepartmentTerms : [],
      defaultLanguageTerms : []
    };
    this.onInit.bind(this);
    this.onInit();
    this.updateProfile.bind(this);

    this.setTarget = element => {
      this._targetButton = element;
    };
  }

  public onInit() :void{
    sp.profiles.myProperties.get().then(user =>{
      console.log(user);
      console.log(React.version);
      let imageUrl = (user.PictureUrl)? user.PictureUrl.replace("MThumb","LThumb") : 
        this.props.context.pageContext.site.absoluteUrl + "/_layouts/15/images/PersonPlaceholder.200x150x32.png";
      let accountName = user.AccountName;
      let title = user.Title;
      //get firstName
      let firstName :string = user.UserProfileProperties.find(item =>{
        return item.Key == "FirstName";
      }).Value;

      let promises = [
        this.setUserTaxonomy(user, "SPS-Location"),
        this.setUserTaxonomy(user, "SPS-Department"),
        this.setUserTaxonomy(user, "SPSNewsLanguage")
      ];
      Promise.all(promises).then(terms =>{
        if(terms[0] == undefined || terms[1] == undefined  || terms[2] == undefined ){
          this.setState({
            showCheckMark : true
          });
        }
        this.setState({
          defaultLocationTerms : terms[0],
          defaultDepartmentTerms : terms[1],
          defaultLanguageTerms : terms[2]
        });
      }, error=>{
        console.log(error);
      });

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
      <Fabric id="_userProfile" >
      <div className={styles.spProfileUpdate}>
        <div className={[styles.login_box, styles.row].join(' ')}>
            <div className={[styles["col-md-12"],styles["col-sm-12"]].join(' ')} style={{textAlign:'center'}}>
                  <div className={styles.line}><h3 style={{padding: '10px', marginTop:0,fontWeight:600}}>{(h % 12) == 0? '12' : h % 12}:{(m < 10 ? '0' + m : m)}:{(s < 10 ? '0' + s : s)} {h < 12 ? 'AM' : 'PM'}</h3></div>
                  <div className={styles.outter}><img src={this.state.imageUrl} className={styles["image-circle"]}/></div>   
                  <h1>{greeting}{this.state.firstName}</h1>
                  <span>{this.state.title}</span>
                  <br/>
                  <div ref={this.setTarget}>
                  <IconButton iconProps={(this.state.open)? { iconName: 'DoubleChevronUp' } : { iconName: 'DoubleChevronDown' }} 
                  onClick={() => this.setState({ open: !this.state.open })}/>
                  </div>
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
                    onChange={this.onTaxChange.bind(this,"Wizdom_Languages")}
                    isTermSetSelectable={false}
                    initialValues={this.state.defaultLanguageTerms}
                  />
                <TaxonomyPicker
                    allowMultipleSelections={false}
                    termsetNameOrID="GeographyHierarchy"
                    panelTitle="Select Location"
                    label="Location"
                    context={this.props.context}
                    onChange={this.onTaxChange.bind(this,"GeographyHierarchy")}
                    isTermSetSelectable={false}
                    initialValues={this.state.defaultLocationTerms}
                  />
                <TaxonomyPicker
                    allowMultipleSelections={false}
                    termsetNameOrID="ApplicableFunction"
                    panelTitle="Select Department"
                    label="Department"
                    context={this.props.context}
                    onChange={this.onTaxChange.bind(this, "ApplicableFunction")}
                    isTermSetSelectable={false}
                    initialValues={this.state.defaultDepartmentTerms}
                  />
                
                <PrimaryButton disabled={this.state.termList.length == 0 } 
                    iconProps={{ iconName: 'AddFriend' }}
                    onClick={this.updateProfile.bind(this)} 
                    style={{float:'right', margin:'20px'}}>Update</PrimaryButton>
                </div>
                <SpProfileReminder target={this._targetButton} isVisible={this.state.showCheckMark}/>
                <br/>
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

  private setUserTaxonomy = async (user: any, key: string) =>{
    let userProp :string = user.UserProfileProperties.find(item =>{
        return item.Key == key;
    }).Value;
    if(userProp){
      return this.getInitialValues(userProp);
    }
  }

  private onTaxChange = (termSetName : string, terms : IPickerTerms,) => {
    let newTermList : IPickerTerms = [];
    if(terms.length !== 0){
      newTermList.push(terms[0]);
    } else {
      if(this.state.termList.length !== 0){
        newTermList = this.state.termList.filter(term =>{
          return term.termSetName == termSetName;
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
      sp.profiles.inBatch(batch).setSingleValueProfileProperty(this.state.accountName,"SPSNewsLanguage",language.name);
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
