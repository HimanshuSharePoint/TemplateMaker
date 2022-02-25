import {
  Version
} from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  escape
} from '@microsoft/sp-lodash-subset';
import styles from './DocumentLibraryWebpartWebPart.module.scss';
import * as strings from 'DocumentLibraryWebpartWebPartStrings';
import * as jQuery from 'jquery';
import 'jqueryui';
import 'kendo-ui-core';
import '@progress/kendo-ui';
import {
  SPComponentLoader
} from '@microsoft/sp-loader';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  result,
  uniqueId
} from 'lodash';
import {
  data
} from 'jquery';

export interface IDocumentLibraryWebpartWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  FileLeafRef: string,
      Author: {
          Title: string
      }
  Created: string,
      Modified: string,
      File_x0020_Type: any,
      Editor: {
          Title: string
      },
      ID: string;
  EncodedAbsUrl: string;
  UniqueId: string;
}

export interface TempLibrarayList {
  UniqueId: string;
  ID: string;
}

export interface TempLibraryLists {
  value: TempLibrarayList[];
}

let KendoDocumentLibraryDataSource;
let DocumentLibraryPath;
let wndApply;
let wndfiles;
let siteUrl;
let fileCount = 0;
let requestHeaders = {
  "accept": "application/json; odata=verbose"
};
let results;
let ResponseLibraryURL;
let serverRelativeUrlToFolder = 'TemplateModule';
let fileInput = jQuery('#getFile');
let def = $.Deferred();
let promises = [];
let fileName;
let serverUrl;

export default class DocumentLibraryWebpartWebPart extends BaseClientSideWebPart < IDocumentLibraryWebpartWebPartProps > {

  public constructor() {
      super();

      SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
      SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
      SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2019.3.1023/styles/kendo.common.min.css')
      SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.1.406/styles/kendo.default-v2.min.css');
      SPComponentLoader.loadCss('https://himanshusharepoint.github.io/TemplateMaker/WinnableKendoCSS.css');


  }

  public render(): void {
      this.domElement.innerHTML = `
      <div class="quick-panel"> 
      <a class="quick-panel-filter" id="ClearFilter"><i class="fa fa-eraser"></i></a>  
   <a class="quick-panel-close" id="CloseID"><i class="fa fa-close"></i></a>  
   <h2 class="spfxttl2">Filters</h2>   
   <div class="SelectSPFX">
     <label for="cars" >Scope:</label>
     <input  id="ScopeKendoDropdown" />
  </div>
  <div class="SelectSPFX">
     <label for="cars">Industry:</label>
     <input  id="IndustryKendoDropdown" />
  </div>
  <div class="SelectSPFX">
     <label for="cars">Month:</label>
     <input  id="MonthKendoDropdown" />
  </div>
  <div class="SelectSPFX">
     <label for="cars">Year:</label>
     <input  id="YearKendoDropdown" />
  </div>
  
  </div>
  <div class="quick-panel-overlay dnone"></div>
  
  <div class="navbar">
  <div class="HeaderTemplate">Template Library</div>
 
 </div>
 <div class="dropdown">
 <button class="dropbtn">New Template
     <i class="fa fa-caret-down"></i>
 </button>
 <div class="dropdown-content">
     <a href="#" id="btnCreateFreshTemplate"> <i class="fa fa-plus"></i>&nbsp;&nbsp;&nbsp; Create fresh template</a>
     <a href="#" id="btnUploadExistingTemplate"><i class="fa fa-upload"></i>&nbsp;&nbsp;&nbsp;Upload existing template</a>
 </div>
</div> 
 <div id="divIconGrid">
  <i title="Filter" class="fa fa-filter" id="IDFilter"></i>   
  <i title="Gridview" class="fa fa-bars"></i>
  <i title="Cardview" class="fa fa-th-large"></i>
  </div>
 <br style="clear:both;"/>
  
    <div class="${ styles.documentLibraryWebpart }">
      
        <div id='KendoGridDocumentLibrary'>
          
        </div>
        <div id="kendoWindow_JobApplication1"></div>
        <div id="kendoWindow_files">
        <input id="getFile" multiple="multiple" type="file" />
        <input id="addFileButton" type="button" value="upload" />
        <div id="TableUpdatedExisting"></div>
        </div>
        <div id="kendoWindow_CreateTemplate"></div>
        
        
      
    </div>`;
    

      this.CreateKendoDataSource();

      $(document).on('click', '#addFileButton', function(e) {
        try {
            alert("Clicking uploading multiple documents into Document Library");
            
            // Get test values from the file input and text input page controls
            uploadnewfile();    
            

        } catch (ex) {
            console.log("Error while Add file into Document Library" + ex.message);
        }
    });
    
   

      

      $(document).on('click','#IDFilter',function(e){
        
        $('.quick-panel').addClass('quick-panel-on');
        $('.quick-panel-overlay').removeClass('dnone');
        $('.quick-panel-overlay').addClass('dblock');
      });

      $(document).on('click',"#CloseID",function(e){
        
        $('.quick-panel').removeClass('quick-panel-on');
        $('.quick-panel-overlay').addClass('dnone');
      });

      var products = $("#ScopeKendoDropdown").kendoDropDownList({
        optionLabel: "Select Scope...",
        dataSource: [ "Cloud Proposal", "Web Proposal", "Analytics Proposal","Platform Engineering Proposal" ]
    }).data("kendoDropDownList");

    var Year = $("#YearKendoDropdown").kendoDropDownList({
        optionLabel: "Select Year...",
        dataSource: [ "2022", "2023", "2024","2025" ]
    }).data("kendoDropDownList");

    var MonthKendo = $("#MonthKendoDropdown").kendoDropDownList({
        optionLabel: "Select Month...",
        dataSource: [ "April", "May", "June","July","August" ]
    }).data("kendoDropDownList");

    var Month = $("#IndustryKendoDropdown").kendoDropDownList({
        optionLabel: "Select Industry...",
        dataSource: [ "IT Servicess", "Health Care", "Pharma","Telecommunication" ]
    }).data("kendoDropDownList");


    
    

      $(document).on('click','#btnUploadExistingTemplate',function(e){
        
        //#region 
        wndfiles = $("#kendoWindow_files").kendoWindow({
            actions: ["Maximize", "Close"],
            width: "auto!important",
            modal: true,
            title: "Upload Existing Template",
            visible: false

        }).data("kendoWindow").center();
        wndfiles.center().open();
        //#endregion
      });

      $(document).on('click', '#btnCreateFreshTemplate', function (e) {
        console.log("Click");
        var randomnumber = Math.floor((Math.random() * 9999999999) + 1);
        var CreateWordFileURL = siteUrl + '/' + "_api/web/GetFolderByServerRelativeUrl('TemplateModule')/Files/Add(url='NewTemplate"+randomnumber+".docx', overwrite=true)";
        $.ajax({
            url: siteUrl + "/_api/contextinfo",
            async: false,
            type: "POST", //While trying to get Context info for user POST is used as type instead of method because body is empty while //request is sent   
            headers: {
                "Accept": "application/json;odata=verbose"
            },
            success: function(contextData) {
                $.ajax({
                    url: CreateWordFileURL,
                    method: 'POST',
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        "X-RequestDigest": contextData.d.GetContextWebInformation.FormDigestValue
                    },

                    success: function(dataitem) {
                        
                               // e.preventDefault();
                              //DocumentLibraryPath = dataItem.EncodedAbsUrl;
                              DocumentLibraryPath = siteUrl + "/_layouts/15/Doc.aspx?sourcedoc=%7B" + dataitem.d.UniqueId + "%7D&file=" + dataitem.d.Name + "&action=default&mobileredirect=true"
                              console.log(DocumentLibraryPath);


                              wndApply = $("#kendoWindow_CreateTemplate").kendoWindow({
                                  actions: ["Maximize", "Close"],  
                                  width: "auto!important",
                                  modal: true,
                                  title: "Create Template",
                                  content: DocumentLibraryPath,
                                  visible: false

                              }).data("kendoWindow").center();
                              wndApply.center().open();
                        
                    },
                    error: function(error) {
                        console.log(JSON.stringify(error));
                    }
                });
            },
            error: function(jqXHR, textStatus, errorThrown) {
                alert('error');
            }
        });


        
      });

    

  }
  
  private onError(error) {
    alert(error.responseText);
    }



  private GetDocumentLibraryData(): Promise < ISPLists > {
      siteUrl = this.context.pageContext.web.absoluteUrl;
      console.log(siteUrl);
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('TemplateModule')/items?$select=*,ID,UniqueId,FileLeafRef,EncodedAbsUrl,File_x0020_Type,Author/Title,Editor/Title&$expand=Author,Editor&$orderby= Created desc", SPHttpClient.configurations.v1)
          .then((Response: SPHttpClientResponse) => {
              return Response.json();
          });

  } 


  

   /// This function will upload multiple documents into Document Library

  private CreateKendoDataSource() {
      try {
          KendoDocumentLibraryDataSource = new kendo.data.DataSource({
              data: [],
              schema: {
                  model: {
                      id: "KendoGridDocumentLibrary",
                      fields: {
                          ID: {
                              editable: false,
                              nullable: false
                          },
                          File_Name: {
                              editable: false,
                              nullable: false
                          },
                          Createdby: {
                              editable: false,
                              nullable: false
                          },
                          Created: {
                              editable: false,
                              nullable: false
                          },
                          LastModifiedOn: {
                              editable: false,
                              nullable: false
                          },
                          LastModifiedBy: {
                              editable: false,
                              nullable: false
                          },
                          FileType: {
                              editable: false,
                              nullable: false
                          },
                          EncodedAbsUrl: {
                              editable: false,
                              nullable: false
                          },
                          UniqueID: {
                              editable: false,
                              nullable: false
                          }
                      }

                  }
              }
          });

          this.GetDocumentLibraryData()
              .then((response) => {
                  this._renderList(response.value);
              });

      } catch (ex) {
          console.log("Error While Create Kendo Data Source" + ex.message);
      }
  } 

  private _renderList(items: ISPList[]): void {
      try {
          items.forEach((item: ISPList) => {
              // Add Data into Kendo Data Source
              KendoDocumentLibraryDataSource.add({
                  File_Name: item.FileLeafRef,
                  Createdby: item.Author.Title,
                  Created: item.Created,
                  LastModifiedOn: item.Modified,
                  LastModifiedBy: item.Editor.Title,
                  filetype: item.File_x0020_Type,
                  ID: item.ID,
                  EncodedAbsUrl: item.EncodedAbsUrl,
                  uniqueId: item.UniqueId
              });

          });

          this.BindKendoGrid();

      } catch (ex) {
          console.log("Error while calling _renderList" + ex.message);
      }

  }





  private BindKendoGrid() {
      try {
          var KendoDocumentLibaryGrid = ( < any > $("#KendoGridDocumentLibrary")).kendoGrid({
              dataSource: {
                  data: KendoDocumentLibraryDataSource.data(),
                  pageSize: 20
              },
              toolbar: ["search"],
              width: "100%",
              scrollable: true,
              resizable: true,
              sortable: true,
              reorderable: true,
              noRecords: true,
              columnMenu: true,
              groupable: true,
              pageable: {
                  refresh: true,
                  pageSizes: true,
                  buttonCount: 5
              },
              filterable: {
                  extra: false,
                  operators: {
                      string: {
                          contains: "Contains"
                      }
                  }
              },
              columns: [{
                      field: "",
                      title: "",
                      width: "50px",
                      template: kendo.template("<a style='cursor:pointer;color:blue;' onclick='return OpenKendoWindow(\#=ID#\);'><span class='fa fa-file-word-o'></span></a>"),
                  },

                  {
                      field: "File_Name",
                      title: "Template title",
                      width: "150px"
                  },
                  {
                      field: "Createdby",
                      title: "Uploaded by",
                      width: "150px"

                  },
                  {
                      field: "Created",
                      title: "Created on",
                      width: "150px",
                      template: "#= kendo.toString(kendo.parseDate(Created, 'yyyy-MM-dd'), 'MM/dd/yyyy') #"

                  },
                  {
                      field: "LastModifiedOn",
                      title: "Modified on",
                      width: "150px",
                      template: "#= kendo.toString(kendo.parseDate(LastModifiedOn, 'yyyy-MM-dd'), 'MM/dd/yyyy') #"

                  },
                  {
                      field: "LastModifiedBy",
                      title: "Modified by",
                      width: "150px"

                  },
                  {
                      command: {
                          text: "Edit template",
                          headerAttributes: {
                              style: "white-space: normal"
                          },
                          click: function(e) {
                              e.preventDefault();
                              var dataItem = this.dataItem($(e.currentTarget).closest("tr"));
                              //DocumentLibraryPath = dataItem.EncodedAbsUrl;
                              DocumentLibraryPath = siteUrl + "/_layouts/15/Doc.aspx?sourcedoc=%7B" + dataItem.uniqueId + "%7D&file=" + dataItem.File_Name + "&action=default&mobileredirect=true"
                              console.log(DocumentLibraryPath);


                              wndApply = $("#kendoWindow_JobApplication1").kendoWindow({
                                  width: "auto!important",
                                  modal: true,
                                  title: "Document",
                                  content: DocumentLibraryPath,
                                  visible: false

                              }).data("kendoWindow").center();
                              wndApply.center().open();

                          }
                      },
                      title: "",
                      field: "",
                      width: 150
                  },
                  {
                      command: {
                          text: "Edit Metadata"
                      },
                      imageClass: "fa fa-trash",
                      title: "",
                      width: 150,
                      hidden: true
                  },
                  {
                      command: {
                          text: "Delete Template"
                      },
                      imageClass: "fa fa-trash",
                      title: "",
                      width: 150
                  },

                  {
                      command: {
                          
                          text: "Create Response",
                          click: function(e) {
                              //#region  Copy Template Document into Temp response Library and Open in Kendo
                              e.preventDefault();
                              var dataItem = this.dataItem($(e.currentTarget).closest("tr"));
                              DocumentLibraryPath = siteUrl + "/_layouts/15/Doc.aspx?sourcedoc=%7B" + dataItem.uniqueId + "%7D&file=" + dataItem.File_Name + "&action=default&mobileredirect=true"
                              console.log(DocumentLibraryPath);
                              var sourceSite = siteUrl + "/_api/web/lists/getByTitle('TemplateModule')/RootFolder/Files('" + dataItem.File_Name + "')/$value"
                              var targetSite = siteUrl + "/_api/web/lists/getByTitle('TempResponselibrary')/RootFolder/Files/Add(url='" + dataItem.File_Name + "',overwrite=true)";
                              var testurl = siteUrl + "/_api/web/folders/GetByUrl('TemplateModule')/Files/getbyurl('" + dataItem.File_Name + "')/copyTo(strNewUrl='TempResponselibrary/" + dataItem.File_Name + "',bOverWrite=true)";
                              console.log(testurl);
                              console.log("Kendo Window URL" + siteUrl + '/TempResponselibrary' + dataItem.File_Name);
                              $.ajax({
                                  url: siteUrl + "/_api/contextinfo",
                                  async: false,
                                  type: "POST", //While trying to get Context info for user POST is used as type instead of method because body is empty while //request is sent   
                                  headers: {
                                      "Accept": "application/json;odata=verbose"
                                  },
                                  success: function(contextData) {
                                      $.ajax({
                                          url: siteUrl + "/_api/web/folders/GetByUrl('TemplateModule')/Files/getbyurl('" + dataItem.File_Name + "')/copyTo(strNewUrl='TempResponselibrary/" + dataItem.File_Name + "',bOverWrite=true)",
                                          method: 'POST',
                                          headers: {
                                              "Accept": "application/json; odata=verbose",
                                              "X-RequestDigest": contextData.d.GetContextWebInformation.FormDigestValue
                                          },

                                          success: function(dataitem) {

                                              

                                              var requestUri = siteUrl + "/_api/web/folders/GetByUrl('TempResponselibrary')/Files/getbyurl('" + dataItem.File_Name + "')";
                                              var array = [];
                                              jQuery.ajax({
                                                  url: requestUri,
                                                  type: "GET",
                                                  dataType: "json",
                                                  headers: requestHeaders,
                                                  async: false,
                                                  success: function(data) {
                                                      //#region  Get file property of Temp response libary from file URL
                                                      if (data.d != undefined) {
                                                          console.log(data.d);
                                                          results = data.d;
                                                      } else {
                                                          results = null;
                                                      }

                                                      if (results.Length > 0) {
                                                          ResponseLibraryURL = siteUrl + "/_layouts/15/Doc.aspx?sourcedoc=%7B" + results.UniqueId + "%7D&file=" + dataItem.File_Name + "&action=default&mobileredirect=true";
                                                      }

                                                      console.log(dataItem);
                                                      wndApply = $("#kendoWindow_JobApplication1").kendoWindow({
                                                          width: "auto!important",
                                                          modal: true,
                                                          title: "Response Document Library",
                                                          content: ResponseLibraryURL,
                                                          visible: false

                                                      }).data("kendoWindow").center();
                                                      wndApply.center().open();
                                                      //#endregion
                                                  },
                                                  error: function(error) {
                                                      console.log(JSON.stringify(error));
                                                      results = null;
                                                  }
                                              });

                                          },
                                          error: function(jqXHR, errorText) {
                                              alert("Problem with copying" + errorText);
                                          }
                                      });
                                  },
                                  error: function(jqXHR, textStatus, errorThrown) {
                                      alert('error');
                                  }
                              });


                              //#endregion 
                          },
                          imageClass: "k-i-info",
                      },
                      title: "",
                      width: 150
                  }

              ]
          }).data("kendoGrid");



      } catch (ex) {
          console.log("Error While Binding Kendo Grid" + ex.message);
      }
  }

 

  protected getFormDigest(webUrl) {
      return $.ajax({
          url: siteUrl + "/_api/contextinfo",
          method: "POST",
          headers: {
              "Accept": "application/json; odata=verbose"
          }
      });
  }

  protected get dataVersion(): Version {
      return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
          pages: [{
              header: {
                  description: strings.PropertyPaneDescription
              },
              groups: [{
                  groupName: strings.BasicGroupName,
                  groupFields: [
                      PropertyPaneTextField('description', {
                          label: strings.DescriptionFieldLabel
                      })
                  ]
              }]
          }]
      };
  }
}
function uploadnewfile() {

    var fileInput = $("#getFile");
    fileName = ( < HTMLInputElement > fileInput[0]).files[0].name;
    fileCount = ( < HTMLInputElement > fileInput[0]).files.length;

    // Get the server URL
    serverUrl = siteUrl;
    var filesUploaded = 0;
    $("#TableUpdatedExisting").append("<h2>Thank you, The files are succesfully uploaded.<br/> Below are the details </h2><br/><table style='width:300px;height:100px;border: 2px solid blue'><tbody><tr><th>File Name</th></tr>");
    for (var i = 0; i < fileCount; i++) {
        // Get the local file as an array buffer
        var getFile = getFileBuffer(i);
        getFile.done(function(arrayBuffer, i) {
            // Add the file to the SharePoint folder.
            var addFile = addFileToFolder(arrayBuffer, i);
           
            addFile.done(function(file, status, xhr) {
                filesUploaded++;
                if (fileCount == filesUploaded) {
                    alert("All files uploaded succesfully");
                    $("#getFiles").val = null;
                    filesUploaded = 0;
                    $("#TableUpdatedExisting").append("</tbody></table>");
                }
            });
            addFile.fail(onerror);
        });
        getFile.fail(onerror);
    }

    function getFileBuffer(i) {

        var deferred = jQuery.Deferred();

        var reader = new FileReader();

        reader.onloadend = function(e) {

            deferred.resolve(e.target.result, i);
        }
        reader.onerror = function(e) {

            deferred.reject(e.target.error);
        }
        reader.readAsArrayBuffer(( < HTMLInputElement > fileInput[0]).files[i]);
        return deferred.promise();

    }
    // Add the file to the file collection in the Shared Documents folder.

    function addFileToFolder(arrayBuffer, i) {

        var index = i;

        // Get the file name from the file input control on the page.

        var fileName = ( < HTMLInputElement > fileInput[0]).files[index].name;

        // Construct the endpoint.
        
        var fileCollectionEndpoint = siteUrl + '/' + "_api/web/GetFolderByServerRelativeUrl('TemplateModule')/Files/Add(url='"+fileName+"', overwrite=true)";
        // Send the request and return the response.
        return jQuery.ajax({
            url: siteUrl + "/_api/contextinfo",
            async: false,
            type: "POST", //While trying to get Context info for user POST is used as type instead of method because body is empty while //request is sent   
            headers: {
                "Accept": "application/json;odata=verbose"
            },
            success: function(contextData) {
                jQuery.ajax({
                    url: fileCollectionEndpoint,
                    type: "POST",
                    data: arrayBuffer,
                    processData: false,
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        "X-RequestDigest": contextData.d.GetContextWebInformation.FormDigestValue
                    },
                    success: function(dataitem) {
                        alert(fileName + "Uploaded succesfully");
                        $("#TableUpdatedExisting").append("<tr><td style='border: 2px solid Red'>"+fileName+"</td></tr>");
                    },
                    error: function(error) {
                        console.log(JSON.stringify(error));
                    }
                });
            },
            error: function(jqXHR, textStatus, errorThrown) {
                alert('error');
            }
        });
        
        // This call returns the SharePoint file.
    }
}
    // Display error messages.
    function onError(error) {
    
    alert(error.responseText);
    }

