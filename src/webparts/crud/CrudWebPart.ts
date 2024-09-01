import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape, times } from '@microsoft/sp-lodash-subset';

import styles from './CrudWebPart.module.scss';
import * as strings from 'CrudWebPartStrings';


import { ISPHttpClientOptions, SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from "@microsoft/sp-http";
import { ISoftwareListItem } from './ISoftwareListItem';

export interface ICrudWebPartProps {
  description: string;
}


export default class CrudWebPart extends BaseClientSideWebPart<ICrudWebPartProps> {



  public render(): void {
    this.domElement.innerHTML = `
      <div>

<div>
  <table border="5" bgcolor="aqua">
    <tr>
      <td>Please Enter Software ID </td>
      <td><input type="text" id="txtID"/></td>
        <td><input type="submit" id="btnRead" value="Read Details"/></td>
    </tr>

    <tr>
      <td>Software Title</td>
      <td><input type="text" id="txtSoftwareTitle"></td>
    </tr>

    <tr>
      <td>Software Name</td>
     <td><input type="text" id="txtSoftwareName"></td>
    </tr>

    <tr>
      <td>Software Vendor</td>
      <td>
        <select name="" id="ddlSoftwareVendor">
          <option value="Sun">sun</option>
          <option value="Oracle">Oracle</option>
          <option value="Microsoft">Microsoft</option>
        </select>
      </td>
    </tr>

    <tr>
      <td>Software Version</td>
      <td><input type="text" id="txtSoftwareVersion"></td>
    </tr>

    <tr>
      <td>Software Description</td>
      <td><textarea name="" id="txtSoftwareDescription" cols="40" rows="5"></textarea></td>
    </tr>

    <tr>
      <td colspan="2" align="center">
      <input type="submit" id="btnSubmit" value="Insert Item"/>
      <input type="submit" id="btnUpdate" value="Update"/>
      <input type="submit" id="btnDelete" value="Delete"/>
      </td>
    </tr>

  </table>
  </div>
  <div id="divStatus"></div>
</div>




      </div>`;
      this._bindEvents();
      this.readAllItems();
  }
  private readAllItems() {

    this._getListItems().then(listItems => {
      let html: string = '<table border="1" width="100%" style= "border-colapse: collapse;">';
      html += '<tr><th>ID</th><th>Title</th><th>Software Name</th><th>Software Vendor</th><th>Software Version</th><th>Software Description</th></tr>';
      listItems.forEach(item => {
        html += `<tr><td>${item.Id}</td><td>${item.Title}</td><td>${item.SoftwareName}</td><td>${item.SoftwareVendor}</td><td>${item.SoftwareVersion}</td><td>${item.SoftwareDescription}</td></tr>`;
      });
      html += '</table>';
      document.getElementById('divStatus').innerHTML = html;
      const listContainer: Element = this.domElement.querySelector('#divStatus');

      listContainer.innerHTML = html;
    })

  }

private _getListItems(): Promise<ISoftwareListItem[]> {
  const url: string = this.context.pageContext.site.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items";
  return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
  .then(response => {
    return response.json();
  })
  .then(json => {
    return json.value;
  }) as Promise<ISoftwareListItem[]>
}


  private _bindEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', ()=> {this.addListItem();});
    this.domElement.querySelector('#btnRead').addEventListener('click', ()=>{this.readListItem();});
       this.domElement.querySelector('#btnUpdate').addEventListener('click', ()=>{this.updateListItem();});
       this.domElement.querySelector('#btnDelete').addEventListener('click', ()=>{this.deleteListItem();});


  }
  private deleteListItem() {

    var title = document.getElementById('txtSoftwareTitle')['value']
    var softwareName = document.getElementById('txtSoftwareName')['value']
    var softwareVersion = document.getElementById('txtSoftwareVersion')['value']
    var softwareVendor = document.getElementById('ddlSoftwareVendor')['value']
    var softwareDescription = document.getElementById('txtSoftwareDescription')['value']

    let id: string= document.getElementById('txtID')['value']

    const url: string = this.context.pageContext.site.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items(" +id+ ")";
    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*",
    };


    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: headers,
    };
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions).then((response: SPHttpClientResponse): void => {
      if (response.status === 204) {
        this.domElement.querySelector('#divStatus').innerHTML = "Item deleted successfully";
      } else {
        this.domElement.querySelector('#divStatus').innerHTML = "Error deleting item";
      }
    });

  }
  private updateListItem() {

    var title = document.getElementById('txtSoftwareTitle')['value']
    var softwareName = document.getElementById('txtSoftwareName')['value']
    var softwareVersion = document.getElementById('txtSoftwareVersion')['value']
    var softwareVendor = document.getElementById('ddlSoftwareVendor')['value']
    var softwareDescription = document.getElementById('txtSoftwareDescription')['value']

    let id: string= document.getElementById('txtID')['value']

    const url: string = this.context.pageContext.site.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items(" +id+ ")";
    const itemBody: any = {
      "Title": title,
      "SoftwareName": softwareName,
      "SoftwareVersion": softwareVersion,
      "SoftwareVendor": softwareVendor,
      "SoftwareDescription": softwareDescription
    };
    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };


    const options: ISPHttpClientOptions = {
      headers: headers,
      body: JSON.stringify(itemBody),
    };
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options).then((response: SPHttpClientResponse): void => {
      if (response.status === 204) {
        this.domElement.querySelector('#divStatus').innerHTML = "Item updated successfully";
      } else {
        this.domElement.querySelector('#divStatus').innerHTML = "Update: Could not update list item. " + response.status + " - " + response.statusText;
      }
    });


  }

  private  readListItem(): void {


    let id: string= document.getElementById('txtID')['value']
   this._getListItemByID(id).then(listItem=>{

    document.getElementById('txtSoftwareTitle')['value']=listItem.Title
    document.getElementById('ddlSoftwareVendor')['value']=listItem.SoftwareVendor
    document.getElementById('txtSoftwareDescription')['value']=listItem.SoftwareDescription
    document.getElementById('txtSoftwareName')['value']=listItem.SoftwareName
    document.getElementById('txtSoftwareVersion')['value']=listItem.SoftwareVersion
   })
.catch((error)=>{
  let message: Element = this.domElement.querySelector('#divStatus');
  message.innerHTML = 'Read: Could not read list item. ' + error;
});
  }

private _getListItemByID(id: string): Promise<ISoftwareListItem> {

  const url: string = this.context.pageContext.site.absoluteUrl+"/_api/web/lists/getbytitle('SoftwareCatalog')/items?$filter=Id eq "+id;
return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
.then((response: SPHttpClientResponse)=> {
  return response.json();
})
.then((listItems: any)=>{
  const untypedItem: any = listItems.value[0]
  const listItem: ISoftwareListItem = untypedItem as ISoftwareListItem
  return listItem
}) as Promise <ISoftwareListItem>
}

  private addListItem(): void {

  var softwaretitle = document.getElementById('txtSoftwareTitle')['value']
  var softwarename = document.getElementById('txtSoftwareName')['value']
  var softwareversion = document.getElementById('txtSoftwareVersion')['value']
  var softwarevendor = document.getElementById('ddlSoftwareVendor')['value']
  var softwareDescription = document.getElementById('txtSoftwareDescription')['value']

const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items";

const itemBody: any = {
  "Title": softwaretitle,
  "SoftwareName": softwarename,
  "SoftwareVersion": softwareversion,
  "SoftwareVendor": softwarevendor,
  "SoftwareDescription": softwareDescription
};




const options: ISPHttpClientOptions = {
  body: JSON.stringify(itemBody),
}

this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, options).then((response: SPHttpClientResponse): void => {

  if (response.status === 201) {
    this.domElement.querySelector('#divStatus').innerHTML = "Item added successfully";

    this.clearForm()


  } else {
    this.domElement.querySelector('#divStatus').innerHTML = "Error while adding item" + response.status + " - "+ response.statusText;
  }
});
}


private clearForm(): void {
  document.getElementById('txtSoftwareTitle')['value']=''
  document.getElementById('ddlSoftwareVendor')['value']=''
  document.getElementById('txtSoftwareDescription')['value']=''
  document.getElementById('txtSoftwareName')['value']=''

}



// protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

