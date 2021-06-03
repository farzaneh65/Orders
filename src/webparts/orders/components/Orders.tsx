import * as React from 'react';
import styles from './Orders.module.scss';
import { IOrdersProps } from './IOrdersProps';
import {IOrdersState} from './IOrdersState'
import { escape } from '@microsoft/sp-lodash-subset';
import {IOrdersList} from './IOrdersList'

import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http'
import { sp } from "@pnp/sp";
import { Web, List, ItemAddResult, PnPClientStorageWrapper } from "sp-pnp-js/lib/pnp"
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";  

import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import {FileTypeIcon, ApplicationType, IconType, ImageSize} from '@pnp/spfx-controls-react/lib/FileTypeIcon'

import{
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
} from 'office-ui-fabric-react'
import { WebInfos } from 'sp-pnp-js/lib/sharepoint/webs';

const textFieldStyles:Partial<ITextFieldStyles>={fieldGroup:{width:500}};
const narrowTextFieldStyles:Partial<ITextFieldStyles>={fieldGroup:{width:100}};
const narrowDropdownStyles: Partial<IDropdownStyles>={dropdown:{width:300}};

export default class Orders extends React.Component<IOrdersProps, IOrdersState> {

  constructor (props: IOrdersProps, state: IOrdersState){
    super(props)
    this.state={
      status:"Ready",
      orderListItem: {
        Id:0,
        Title:"",
        Quantity: 0,
        DestinationId:1,
        Description:"",
        LinkToFile:"",
        Owners:[]
      },
      filePickerResult:null,
      Users:[],
      message:""
    }
  }

 

  ////////// Get Destinations ////////////////
  private _getDestinations(): Promise<any[]>{
    const url: string =this.props.siteURL+"/_api/web/lists/getbytitle('Destinations')/Items?$filter=Active eq 1"    
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then
      (response =>{
        return response.json()
      })
    .then
      (json=>{
        return json.value
      }) as Promise<any[]>
      
  }


 ////////// Get Orders-For debugging ////////////////
  private _getOrderss(): Promise<any[]>{
    const url: string =this.props.siteURL+"/_api/web/lists/getbytitle('Orders')/Items"   
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then(response =>{
        return response.json()
      }).then
      (json=>{
        console.log(json.value)
        return json.value
      }) as Promise<any[]>
  }

////////// Upload Recieved File To Document Library ////////////////  
  private  async uploadFileToLibrary(ItemId)
  {
    const web: Web = new Web(this.props.siteURL);
    web.folders.getByName('OrdersFiles').folders.add(`Order-${ItemId}`)
    .then(async fl=>{
      web.getFolderByServerRelativeUrl(`OrdersFiles/Order-${ItemId}`).files
      .add(this.state.filePickerResult[0].fileName,await this.state.filePickerResult[0].            downloadFileContent(), true )
      .then((result)=>{
        console.log(result)
        return true;
    }) as Promise<boolean>
    // .catch(err=>{
    //   console.log(err.errorMessage)
    //   return false
    // })
    })
    .catch(error=>{
      console.log(error.errorMessage)
      return false
    })as Promise<boolean>
  }
////////// Component Did Mount ////////////////
  private options: any[] = [];
  public async componentDidMount(): Promise<void>{
    //this._getOrderss()
    var destinations=await this._getDestinations()
    console.log(this)
    destinations.forEach(dest => {
      this.options.push({
        key:dest.ID,
        text:dest.Title
      })
    });
  }

////////// Add Item To the Order List ////////////////  
  @autobind
  public btnAdd_click():void {
    this.state.Users.forEach(element => {
      console.log(element)
    });
    const web: Web = new Web(this.props.siteURL);
    web.lists.getById('B1C37CDB-2E4D-41F5-8381-B22701ABA693').items
    .add({
      Title:this.state.orderListItem.Title,
      Quantity: this.state.orderListItem.Quantity,
      DestinationId:this.state.orderListItem.DestinationId,
      Description:this.state.orderListItem.Description,
      OwnersId:{
        "results": this.state.Users
      }
    })
    .then(async r=>{
      console.log(r)
      r.item.attachmentFiles.add(this.state.filePickerResult[0].fileName,await this.state.filePickerResult[0].downloadFileContent())
      console.log(r.data.Id)
      if(this.uploadFileToLibrary(r.data.Id)){
        var fileUrl=this.props.siteURL+`/OrdersFiles/Order-${r.data.Id}/${this.state.filePickerResult[0].fileName}`
        console.log(fileUrl)
        r.item.update({
          LinkToFile: {
            "__metadata": { type: "SP.FieldUrlValue" },
            Description: this.state.filePickerResult[0].fileNameWithoutExtension,
            Url: fileUrl
        }
          
        })
        .then((i=>{
          alert('Order is registered successfully.')
        }))
        .catch(err=>{
          console.log(err.errorMessage)
        })
      }
      
    })
    .catch(error=>{
      console.log(error.errorMessage)
    })
  }

  @autobind
  private _getPeoplePickerItems(items: any[]) {
    ///console.log('Items:', items);
    let getSelectedUsers=[];  
    for (let item in items) {  
      getSelectedUsers.push(items[item].id);  
    }  
    this.setState({ Users: getSelectedUsers });
    
  }
  

  public render(): React.ReactElement<IOrdersProps> {
    const dropdownRef=React.createRef<IDropdown>()
    return (
      <div className={ styles.orders }>
          <TextField 
            label='Item'
            required={true}
            value={(this.state.orderListItem.Title).toString()}
            styles={textFieldStyles}
            onChanged={e=>{this.state.orderListItem.Title=e;}}
          />
          <TextField 
            label='Quantity'
            required={true}
            type='number'
            //value={(this.state.orderListItem.Quantity)}
            styles={textFieldStyles}
            onChanged={e=>{this.state.orderListItem.Quantity=e;}}
          />
           <Dropdown 
            componentRef={dropdownRef}
            placeholder="Select an option"
            label='Destination'
            options={this.options}
            defaultSelectedKey=""
            required
            styles={narrowDropdownStyles}
            onChanged={e=>{this.state.orderListItem.DestinationId=Number(e.key);}}
          />
           
        <PeoplePicker  
          context={this.props.context as any}  
          titleText="Owners"  
          personSelectionLimit={3}  
          showtooltip={true}  
          required={true}  
          disabled={false}  
          //selectedItems={this._getPeoplePickerItems} 
           
         onChange={this._getPeoplePickerItems}
          showHiddenInUI={false}  
          ensureUser={true}  
          principalTypes={[PrincipalType.User]}  
          resolveDelay={1000} />  
        
          <TextField 
            label='Description' 
            type='string'
            //value={(this.state.orderListItem.Quantity)}
            styles={textFieldStyles}
            onChanged={e=>{this.state.orderListItem.Quantity=e;}}
          />

          <p></p>
        <FilePicker    
          buttonLabel="Select File"    
          
          onSave={(filePickerResult: IFilePickerResult[]) => { this.setState({filePickerResult});   //alert('File has been uploaded successfully.');
          document.getElementById('msg').innerHTML=filePickerResult[0].fileName+' has been uploaded successfully!!' 
          console.log(this.state.filePickerResult)
  
          }}    
        onChange={(filePickerResult: IFilePickerResult[]) => { this.setState({filePickerResult});//alert('File has been uploaded successfully!!');
        document.getElementById('msg').innerHTML=filePickerResult[0].fileName+' has been uploaded successfully!!'
        console.log(this.state.filePickerResult)
  
          }}    
          context={this.props.context as any}
        /> 
        <p id="msg" style={{color:'green'}}></p>
        <p className={styles.title}>
           <PrimaryButton 
            style={{backgroundColor:'rgb(34, 122, 138)'}}
            text='Add'
            title='Add'
            onClick={this.btnAdd_click}
          />
        </p>  
      </div>
     
    );        
  }
}


