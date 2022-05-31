import * as React from 'react';
import styles from './CrudWithReact.module.scss';
import { ICrudWithReactProps } from './ICrudWithReactProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { ISoftwareListItem } from './ISoftwareListItem';
import { ICrudWithReactState } from './ICrudWithReactState';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';


import {
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  Dropdown,
  IDropdown,
  IDropdownOption,
  ITextFieldStyles,
  IDropdownStyles,
  DetailsRowCheck,
  Selection
} from 'office-ui-fabric-react';


  // Configure the columns for the DetailsList component
  let _softwareListColumns = [
    {
      key: 'ID',
      name: 'ID',
      fieldName: 'ID',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'Title',
      name: 'Title',
      fieldName: 'Title',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'SoftwareName',
      name: 'SoftwareName',
      fieldName: 'SoftwareName',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'SoftwareVendor',
      name: 'SoftwareVendor',
      fieldName: 'SoftwareVendor',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'SoftwareVersion',
      name: 'SoftwareVersion',
      fieldName: 'SoftwareVersion',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'SoftwareDescription',
      name: 'SoftwareDescription',
      fieldName: 'SoftwareDescription',
      minWidth: 50,
      maxWidth: 150,
      isResizable: true
    }  
  ];
  

  const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };
  const narrowTextFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 100 } };
  const narrowDropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

export default class CrudWithReact extends React.Component<ICrudWithReactProps, ICrudWithReactState> {

  
  private _selection: Selection;
  
  
  private _onItemsSelectionChanged = () => {
    
    
    this.setState({
      SoftwareListItem: (this._selection.getSelection()[0] as ISoftwareListItem)
    });
  }

  constructor(props: ICrudWithReactProps, state: ICrudWithReactState) {
    super(props);

    this.state = {
      status: 'Ready',
      SoftwareListItems: [],
      SoftwareListItem: {
        Id: 0,
        Title: "",
        SoftwareName: "",
        SoftwareDescription: "",
        SoftwareVendor: "Select an option",
        SoftwareVersion: ""
      }
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged,
    }); 

    
  }


  private _getListItems(): Promise<ISoftwareListItem[]> {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicrosoftSoftware')/items";
    return this.props.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then(response => {
    return response.json();
    })
    .then(json => {
    return json.value;
    }) as Promise<ISoftwareListItem[]>;
    }

    public bindDetailsList(message: string) : void {

      this._getListItems().then(listItems => {
        this.setState({ SoftwareListItems: listItems,status: message});
      });
    }
  
    public componentDidMount(): void {
      this.bindDetailsList("All Records have been loaded Successfully");  
  
      
    }



    @autobind
  public btnAdd_click(): void {  

    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicrosoftSoftware')/items";          
       
       const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(this.state.SoftwareListItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
     
      if (response.status === 201) {
        this.bindDetailsList("Record added and All Records were loaded Successfully");         

       
       
      } else {
        let errormessage: string = "An error has occured i.e.  " + response.status + " - " + response.statusText;
        this.setState({status: errormessage});        
      }
    });
  }


  @autobind
  public btnUpdate_click(): void {

    let id: number = this.state.SoftwareListItem.Id;

    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicrosoftSoftware')/items(" + id + ")";          
      
    
    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };
       
       const spHttpClientOptions: ISPHttpClientOptions = {
        "headers": headers,
        "body": JSON.stringify(this.state.SoftwareListItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
     
      if (response.status === 204) {
        this.bindDetailsList("Record Updated and All Records were loaded Successfully");                
       
      } else {
        let errormessage: string = "An error has occured i.e.  " + response.status + " - " + response.statusText;
        this.setState({status: errormessage});        
      }
    });
  }


  @autobind
  public btnDelete_click(): void {
    let id: number = this.state.SoftwareListItem.Id;

    

    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('MicrosoftSoftware')/items(" + id + ")";          

    
    const headers: any = { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers
    };


    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if (response.status === 204) {
        alert("record got deleted successfully....");
        this.bindDetailsList("Record deleted and All Records were loaded Successfully");   
        
      } else {
        let errormessage: string = "An error has occured i.e.  " + response.status + " - " + response.statusText;
        this.setState({status: errormessage}); 
      }
    });
  }


  public render(): React.ReactElement<ICrudWithReactProps> {
    
    const dropdownRef = React.createRef<IDropdown>();

    return (
      <div className={ styles.crudWithReact }>
        
        <TextField                  
                  label="ID"
                  required={ false } 
                  value={ (this.state.SoftwareListItem.Id).toString()}
                  styles={textFieldStyles}
                  onChanged={e => {this.state.SoftwareListItem.Id=e;}}
                />
                <TextField                  
                  label="Software Title"
                  required={ true } 
                  value={ (this.state.SoftwareListItem.Title)}
                  styles={textFieldStyles}
                  onChanged={e => {this.state.SoftwareListItem.Title=e;}}
                />
                <TextField                  
                  label="Software Name"
                  required={ true } 
                  value={ (this.state.SoftwareListItem.SoftwareName)}
                  styles={textFieldStyles}
                  onChanged={e => {this.state.SoftwareListItem.SoftwareName=e;}}
                />
                <TextField                  
                  label="Software Description"
                  required={ true } 
                  value={ (this.state.SoftwareListItem.SoftwareDescription)}
                  styles={textFieldStyles}
                  onChanged={e => {this.state.SoftwareListItem.SoftwareDescription=e;}}
                />
                <TextField                  
                  label="Software Version"
                  required={ true } 
                  value={ (this.state.SoftwareListItem.SoftwareVersion)}
                  styles={textFieldStyles}
                  onChanged={e => {this.state.SoftwareListItem.SoftwareVersion=e;}}
                />
                <Dropdown 
                componentRef={dropdownRef}                
                placeholder="Select an option"
                label="Software Vendor"
                options={[
                  { key: 'Microsoft', text: 'Microsoft'},
                  { key: 'Sun', text: 'Sun' },
                  { key: 'Oracle', text: 'Oracle'},
                  { key: 'Google', text: 'Google'}
                  
                ]}
                defaultSelectedKey={this.state.SoftwareListItem.SoftwareVendor}
                required        
                styles={narrowDropdownStyles}        
                onChanged={e => {this.state.SoftwareListItem.SoftwareVendor=e.text;}}
                />




<p className={styles.title}>
                   <PrimaryButton
                    text='Add'      
                    title='Add'              
                    onClick={this.btnAdd_click}
                  />

                  <PrimaryButton
                    text='Update'                    
                    onClick={this.btnUpdate_click}
                  />

                  <PrimaryButton
                    text='Delete'                    
                    onClick={this.btnDelete_click}
                  />
                </p> 


                <div id="divStatus">
                  {this.state.status}
                </div>

                <div>
                <DetailsList
                      items={ this.state.SoftwareListItems}
                      columns={ _softwareListColumns }
                      setKey='Id'
                      checkboxVisibility={ CheckboxVisibility.onHover}
                      selectionMode={ SelectionMode.single}
                      layoutMode={ DetailsListLayoutMode.fixedColumns }
                      compact={ true }
                      selection={this._selection}                                         
                  />
                  </div>  


      </div>
    );


  }
}
