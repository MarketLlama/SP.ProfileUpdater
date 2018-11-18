import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

export interface ISpProfileUpdateState{
    time : Date;
    firstName : string;
    imageUrl : string;
    title : string;
    accountName : string;
    termList : IPickerTerms;
    showCheckMark : boolean;
    open : boolean;
    defaultLocationTerms : IPickerTerms;
    defaultLanguageTerms : IPickerTerms;
    defaultDepartmentTerms : IPickerTerms;
  }