import * as React from 'react';
import styles from './ConsumeRestApiwp.module.scss';
import { IConsumeRestApiwpProps } from './IConsumeRestApiwpProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { ICustomer} from './ICustomer';
import { ICustomerState} from './ICustomerState';


import {
  AadHttpClient,
  AadHttpClientFactory,
  HttpClientResponse
} from '@microsoft/sp-http';


import {
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode, 
  ITextFieldStyles,
  IDropdownStyles,
  DetailsRowCheck,
  Selection
} from 'office-ui-fabric-react';

let _customerListColumns = [
  {
    key: 'id',
    name: 'id',
    fieldName: 'id',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'joined',
    name: 'joined',
    fieldName: 'joined',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'name',
    name: 'name',
    fieldName: 'name',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'city',
    name: 'city',
    fieldName: 'city',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'orderTotal',
    name: 'orderTotal',
    fieldName: 'orderTotal',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  }
];

const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };
  const narrowTextFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 100 } };

  const siteUrl: string ='https://mynewsecuredfunctionappdemo.azurewebsites.net';

export default class ConsumeRestApiwp extends React.Component<IConsumeRestApiwpProps, ICustomerState> {

  private _selection: Selection;

  private _onItemsSelectionChanged = () => {
    
    
    this.setState({
      CustomerListItem: (this._selection.getSelection()[0] as ICustomer)
    });
  }

  constructor(props: IConsumeRestApiwpProps, state: ICustomerState) {
    super(props);
    this.state = {
      status: 'Ready',
      CustomerListItems: [],
      CustomerListItem: {
        id: "",
        joined: "",
        name: "",
        city: "",
        orderTotal: 0
        
      }
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged,
    }); 

  }

  public _getListItems(): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      this.props.context.aadHttpClientFactory
        .getClient(siteUrl)
        .then((client: AadHttpClient) => {
          const azurefunctionendpoint: string = `${siteUrl}/api/customerdetails`;
          client.get(azurefunctionendpoint, AadHttpClient.configurations.v1)
          .then((response: HttpClientResponse) => {
            return response.json();
          })
          .then((jsonResponse: ICustomer[]) => {
            resolve(jsonResponse);
          })
          .catch((error) => {
            reject(error);
          });
        });
    });
  }

  



  public bindDetailsList(message: string) : void {

    this._getListItems().then(listItems => {
      console.log(listItems);
      this.setState({ CustomerListItems: listItems,status: message});
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList("All Records have been loaded Successfully");  

    
  }
  public render(): React.ReactElement<IConsumeRestApiwpProps> {
    return (
      <div className={ styles.consumeRestApiwp }>
       

          
                <DetailsList
                      items={ this.state.CustomerListItems}
                      columns={ _customerListColumns }
                      setKey='id'
                      checkboxVisibility={ CheckboxVisibility.onHover}
                      selectionMode={ SelectionMode.single}
                      layoutMode={ DetailsListLayoutMode.fixedColumns }
                      compact={ false }
                      selection={this._selection}                   
                  />
                
              
            </div>
       
     
    );
  }
}
