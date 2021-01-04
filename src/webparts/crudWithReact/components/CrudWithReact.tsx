import * as React from 'react';
import styles from './CrudWithReact.module.scss';
import { ICrudWithReactProps } from './ICrudWithReactProps';
import { escape } from '@microsoft/sp-lodash-subset';
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
  DetailsRowCheck,
  Selection,
  IDropdownStyles
} from 'office-ui-fabric-react';
import { IListItem } from './IListItem';

let _carListColumns = [
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
    key: 'Model',
    name: 'Model',
    fieldName: 'Model',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'Fuel',
    name: 'Fuel',
    fieldName: 'Fuel',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  }
];

const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };
const narrowTextFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };
const narrowDropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

export default class CrudWithReact extends React.Component<ICrudWithReactProps, ICrudWithReactState> {

  private _selection: Selection;

  private _onItemSelectionChanged = () => {
    this.setState({
      ListItem: (this._selection.getSelection()[0] as IListItem)
    });
  }

  constructor(props: ICrudWithReactProps, state: ICrudWithReactState) {
    super(props);

    this.state = {
      status: 'Ready',
      ListItems: [],
      ListItem: {
        Id: 0,
        Title: "",
        Model: "",
        Fuel: "Select an option"
      }
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemSelectionChanged
    });
  }

  private _getListItems(): Promise<IListItem[]> {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('CarInventory')/items";
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    })
    .then(json => {
      return json.value;
    }) as Promise<IListItem[]>;
  }

  public bindDetailsList(message: string) : void {
    this._getListItems().then(listItems => {
      this.setState({ ListItems: listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList("All records have been loaded successfully");
  }

  @autobind
  public btnAdd_click(): void {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('CarInventory')/items";

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(this.state.ListItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 201) {
          this.bindDetailsList("Record added and all records loaded successfully");
        }
        else {
          let errorMessage: string = "An error has occured: " + response.status + " - " + response.statusText;
          this.setState({status: errorMessage});
        }
      });
  }

  @autobind
  public btnUpdate_click(): void {
    let id: number = this.state.ListItem.Id;
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('CarInventory')/items(" + id + ")";

    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*"
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body": JSON.stringify(this.state.ListItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 204) {
          this.bindDetailsList("Record updated and all records loaded successfully");
        }
        else {
          let errorMessage: string = "An error has occured: " + response.status + " - " + response.statusText;
          this.setState({status: errorMessage});
        }
      });
  }

  @autobind
  public btnDelete_click(): void {
    let id: number = this.state.ListItem.Id;
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('CarInventory')/items(" + id + ")";

    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*"
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 204) {
          alert("Record deleted successfully");
          this.bindDetailsList("Record deleted and all records loaded successfully");
        }
        else {
          let errorMessage: string = "An error has occured: " + response.status + " - " + response.statusText;
          this.setState({status: errorMessage});
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
          value={ (this.state.ListItem.Id).toString() }
          styles={ textFieldStyles }
          onChanged={ e => { this.state.ListItem.Id=e }}
        />
        <TextField 
          label="Title"
          required={ false }
          value={ (this.state.ListItem.Title).toString() }
          styles={ textFieldStyles }
          onChanged={ e=>{this.state.ListItem.Title=e }}
        />
        <TextField 
          label="Model"
          required={ false }
          value={ (this.state.ListItem.Model).toString() }
          styles={ textFieldStyles }
          onChanged={ e=>{this.state.ListItem.Model=e }}
        />
        <Dropdown 
          componentRef={dropdownRef}
          placeholder="Select an option"
          label="Fuel"
          options={[
            { key: 'Petrol', text: 'Petrol' },
            { key: 'Hybrid', text: 'Hybrid' },
            { key: 'Diesel', text: 'Diesel' }
          ]}
          defaultSelectedKey={this.state.ListItem.Fuel}
          required
          styles={narrowDropdownStyles}
          onChanged={ e=> { this.state.ListItem.Fuel=e.text }}
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
            items={ this.state.ListItems }
            columns={_carListColumns}
            setKey='Id'
            checkboxVisibility={CheckboxVisibility.always}
            selectionMode={SelectionMode.single}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            compact={true}
            selection={this._selection}
          />
        </div>
      </div>
    );
  }
}
