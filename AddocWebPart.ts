import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
require("@pnp/logging");
require("@pnp/common");
require("@pnp/odata");
import { Item, sp, Web } from "@pnp/sp/presets/all";
import * as moment from "moment";
import { escape } from '@microsoft/sp-lodash-subset';
import * as jQuery from 'jquery';
import * as bootstrap from 'bootstrap';
require('bootstrap');
import { taxonomy, ITerm, ITermSet, ITermStore } from '@pnp/sp-taxonomy';
import styles from './AddocWebPart.module.scss';
//require('styles');
import * as strings from 'AddocWebPartStrings';
import {

  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration,
  ISPHttpClientOptions

} from '@microsoft/sp-http';
import { ConsoleListener } from '@pnp/logging';


export interface IAddocWebPartProps {
  description: string;
  RestrictedFields:string;
  UnsupportedPreviewFileTypes:string;
  TermGroupID:string;
  TermSetID:string;
  TermDocTypeID:string;
  TermCollection:string;
  ObjectStores: string;
  AllDocumentTypes:string;
}
export interface IPTerm {
  parent?: string;
  id: string;
  name: string;
}

export interface ITaxonomyPopulatorState {
  terms: IPTerm[];
}

let that: any;
let uploadFile: File;
let metadataArray:{internalName:string,value:any, required:any, fieldtype:any}[]=[];

let updateProperties = {

};
let doclib:any;
let termName:any;
let termExpression:any;
let ctvalue:any;

var AllOUClasses=[];
var RegionalOUClasses=[];

var AllConfigFields=[];
var ObjstoreConfigFields=[];
var CTConfigFields=[];
var DocumentTypeConfigFields=[];
var AllDoumentTypes=[];
var ObjstoreArr=[];
var CTDocTypes=[];
var AllTerms=[];
var DistinctTerms=[];

export default class AddocWebPart extends BaseClientSideWebPart<IAddocWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
    sp.setup({
    spfxContext: this.context
    });

    });
    }
  public render(): void {
    let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(cssURL);
    //SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js',  { globalExportsName: 'jQuery' }).then((): void => {});
    this.domElement.innerHTML = `
        <div class="container">
          <div class="row">
            <div class="col-sm-3 col-md-3 col-lg-3">
              <label style="font-weight: 400"> Select an Object Store</label>
              <select name="region" id="region" tabindex="0">
                <option value="defaultOS">--Select--</option>

              </select>
            </div>
            <div class="col-sm-3 col-md-3 col-lg-3">
              <label style="font-weight: 400"> Select Content Types</label>
              <select name="contenttypes" id="contenttypes"  style="width: 126px;" >
                <option value="defaultCT">--Select--</option>
              </select>
            </div>
            <div class="col-sm-6 col-md-6 col-lg-6">
            <table>
            <tr>
              <td>
                <label style="font-weight: 400"> Select Document Type</label>
                <input type="text" id="txtFilter" name="txtFilter" class="txtFilter" placeholder="Search.." onkeyup="filterFunction()" />
                <select name="terms" id="terms"  style="width: 140px;" >
                  <option value="defaultTerms">--Select--</option>
                </select>
              </td>
              </tr>
              <tr id="trMultiCommit" style="display:none;">
              <td>
                <input type="checkbox" id="chkMulticommit" name="chkMulticommit" value="MultiCommit">
                <label for="MultiCommit"> Multi Commit</label>
              </td>
              </tr>
              </table>
            </div>
          </div>
          <div class="row" style="margin-top: 1%;border-top: 1px solid black;padding-top: 1%;">
            <div id="ctfields">
            </div>
          </div>
          <div class="row">
            <form style="margin-left: 35%;margin-top: 2%;">
              <input type="button" id="Save" value="Save"
                style="display:none; background-color: #f7957ff2;  font-weight: 900;" >

              <input type="file" id="inputFile" name="filename" style="display:inline" >

              <input type="button" id="submit" value="Upload"
                          style="background-color: #90efcc;
              font-weight: 900;">

              <input type="button" id="Clear" value="Clear"
                          style="background-color: #a9b8dc;
              font-weight: 900;">
            </form>
          </div>
          <div id="myProgress" style="display:none;
          width: 100%;
          background-color: #d01515;">
              <div id="myBar" style=" width: 50%;
              height: 35px;
              background-color: green;
              color: #ffffff !important;
              font-weight: 900;
              padding-left: 12%;
              padding-top: 7px;" ><span>Uploading the file .. Please Wait !!!</span></div>

            </div>
          </div>
          <BR\><BR\>
          <div style="overflow-x: auto !important;"> <table id="tblMultiCommit" className="tblMultiCommit" style="border: 2px solid #b14419 !important; width:100%;"> </table> </div>
          <div class="row" style="margin-top: 2%;">
            <iframe id="previewImg" style="width:0px;height:0px"/>
          </div>
        </div> `;

        if(! this.properties.ObjectStores)
        {
          this.properties.ObjectStores="FNAMERICA";
        }

        if(!this.properties.TermCollection)
        {
          this.properties.TermCollection=
          `SensitiveInformationlabel|174c69c4-0236-4bee-99d2-6e1187578b09\nKeysightBusiness|ff67a40f-9fe3-4282-9c29-c8a4b1c9e559\nMarkingSet|b713949e-0595-4f76-99ad-b2e23bb40b0e\nEnvistaCountryName|41eefefc-f597-4490-b6c0-e0e97970e9c5\nEnvistaDocumentType|cfc2ac44-5f83-4564-b5ae-ea6181275647\nReportName1|aa7455d4-1dad-422e-8d0d-49837e26ab60`;
        }

        if(! this.properties.AllDocumentTypes)
        {
          this.properties.AllDocumentTypes="Yes";
        }

      this.PopulateObjectStores();

      this.GetObjectLists().then((response) => {
        response.value.forEach((opt: { Title: any;   }) => {
            ObjstoreArr.push(opt.Title.toUpperCase().trim());
        });
      });

      if(!this.properties.TermGroupID)
      {
        this.properties.TermGroupID="0b1d38f8-f44e-4d13-8b16-c2f308b60021";
      }

      if(!this.properties.TermSetID)
      {
        this.properties.TermSetID="3fac10fe-90d7-49f2-8828-502e31f21174";
      }

      if(!this.properties.TermDocTypeID)
      {
        this.properties.TermDocTypeID="125db7d9-35d6-473a-b123-d67834dac25d";
      }

      this.getterms();
      this.callJQuery();
      this.GetAllConfigFields();
      this.GetAllOUClasses();
      if(!this.properties.RestrictedFields)
      {
        this.properties.RestrictedFields="Test";
      }

      if(!this.properties.UnsupportedPreviewFileTypes)
      {
        this.properties.UnsupportedPreviewFileTypes="zip";
      }
  }

  private callJQuery()
  {
    that=this;

    jQuery('#region').change(function () {
      let selvalue = jQuery('#region').val();
      jQuery('#terms').empty();
      jQuery('#terms').append(new Option("--Select--", "defaultTerms"));
      if(selvalue != "defaultOS" )
      {
        AllTerms.sort(function (x, y) {
          let a = x[0].toUpperCase(),
              b = y[0].toUpperCase();
          return a == b ? 0 : a > b ? 1 : -1;
        });

        for(var kk=0;kk<AllTerms.length;kk++)
        {
          jQuery('#terms').append(new Option(AllTerms[kk][0], AllTerms[kk][1]));
        }
        setTimeout(that.ManageDocumentTypeDisplayNameKGS,2000);
        that.ManageCTDocTypes(selvalue);
      }

      doclib=selvalue;
      that.callContentTypeAPI(selvalue);
      that.GetObjectStoreOUClasses(selvalue);
      that.GetObjectStoreConfigFields(selvalue);
      //that.ManageCTDocTypes(selvalue);
    });

    jQuery('#terms').change(function () {
      let termvalue = jQuery('#terms').val();
      termExpression=termvalue.toString();
      termName=(termvalue.toString()).split("|")[0];
      that.GetDocumentTypeFields(termName);
      //that.ManageMandatoryFields();
      that.ManageDocumentTypeFields();
      //updateProperties.TermDocumentType=termExpression;
      if(termvalue != "defaultTerms")
      {
          $("#trMultiCommit").show();
      }
      else
      {
        $("#trMultiCommit").hide();
      }
    });

    jQuery('#contenttypes').change(function () {
      ctvalue = jQuery('#contenttypes').val();
      let listvalue = jQuery('#region').val();
      that.callContentTypeFieldsAPI(ctvalue,listvalue);
      //that.PopulateOUClasses();
      var CTNAme=$("#contenttypes option[value='"+ctvalue+"']")[0].innerText;
      that.GetCTConfigFields(CTNAme);
    });

    jQuery('#inputFile').change(function () {
      uploadFile=(document.getElementById("inputFile") as HTMLInputElement).files[0];
      if(uploadFile)
      {
        var strFileNameCol=uploadFile.name.split('.');
        var fileExten=strFileNameCol[strFileNameCol.length-1];
        let RestrictedFileTypes=(that.properties.UnsupportedPreviewFileTypes).split('\n');

        if(RestrictedFileTypes.indexOf(fileExten.toLowerCase())  == -1)
        {
          jQuery("#previewImg").attr("src", URL.createObjectURL(uploadFile));
          jQuery("#previewImg").css("width","100%");
          jQuery("#previewImg").css("height","500px");
        }
        else
        {
          jQuery("#previewImg").attr("src", "");
          jQuery("#previewImg").css("width","100%");
          jQuery("#previewImg").css("height","2px");
         // alert("Preview not available for this file");
        }
      }

    });

    jQuery("#submit").click(function(){
      $("#myProgress").show();
      updateProperties={};
      let reqCheck=0;
      var elementMultiCommitChk = <HTMLInputElement> document.getElementById("chkMulticommit");

      if(elementMultiCommitChk.checked == false)
      {
          reqCheck=that.ValidateInputs();
          while(metadataArray.length>0)
          {
            metadataArray.pop();
          }
          that.fetchMetadata();
          metadataArray.forEach(met=>{
          if(met.required=="true" && (met.value==null||met.value==""))
          {
            alert("Please fill all the required field");
            reqCheck=1;
          }
        });
      }
      else
      {
        var elementTableMC= <HTMLElement> document.getElementById("tblMultiCommit");
        var mcRows = elementTableMC.getElementsByTagName("tr");
        if(mcRows.length<2)
        {
          alert("Please enter meta data for Multi-commit");
          reqCheck=1;
        }
      }

      uploadFile=(document.getElementById("inputFile") as HTMLInputElement).files[0];
      if(!uploadFile)
      {
        alert("Please select file to be uploaded ..");
        reqCheck=1;
      }

      if(reqCheck==0)
      {
        uploadFile=(document.getElementById("inputFile") as HTMLInputElement).files[0];
        if(elementMultiCommitChk.checked == false)
        {
          that.uploadFiles();
        }
        else
        {
          that.InitiateMultiUpload();
          // Call Function that will read each row and call Upload Function in loop
        }
      }
      else
      {
        $("#myProgress").hide();
      }

    });

    jQuery("#Save").click(function(){
      let reqCheck=0;

      var elementTableMC= <HTMLElement> document.getElementById("tblMultiCommit");
      var mcRows = elementTableMC.getElementsByTagName("tr");
      if(mcRows.length>10)
      {
        reqCheck=1;
        alert("You have exceeded maximum limit of Orders/Records for Multi-Commit");
      }
      else
      {
        reqCheck=that.ValidateInputs();
      }

      if(reqCheck == 0)
      {
          that.AddRecord();
          alert("Record to Added to Multi-Commit List");
      }
    });

    jQuery("#Clear").click(function(){
      var inpFieldColnTBC=$(".inputField");
      for(var f=0;f<inpFieldColnTBC.length;f++)
      {
        var 	strFieldID=inpFieldColnTBC[f].id;
        if(!((strFieldID == "OUClassName")    || (strFieldID == "Date_x0020_Filed") || (strFieldID == "Document_x0020_Date")))
        {
          $("#"+strFieldID).val("");
        }
      }

      $("#inputFile").val("");
      $("#previewImg").attr("src","");
      $("#previewImg").css("height","1px");

    });

    jQuery('#chkMulticommit').change(function () {
      var elementMultiCommitChk = <HTMLInputElement> document.getElementById("chkMulticommit");
      var elementTermsSelection= <HTMLInputElement> document.getElementById("terms");
      var elementCTSelection= <HTMLInputElement> document.getElementById("contenttypes");
      var elementOSSelection= <HTMLInputElement> document.getElementById("region");
      var elementFilterTxt= <HTMLInputElement> document.getElementById("txtFilter");
      //var elementOUClass= <HTMLInputElement> document.getElementById("OUClassName");

      if(elementMultiCommitChk.checked == true )
      {
        $("#Save").show();
        elementTermsSelection.disabled=true;
        elementCTSelection.disabled=true;
        elementOSSelection.disabled=true;
        elementFilterTxt.disabled=true;
        //elementOUClass.disabled=true;
        that.CreateDefaultTableStructure();
      }
      else
      {
        $("#Save").hide();
        elementTermsSelection.disabled=false;
        elementCTSelection.disabled=false;
        elementOSSelection.disabled=false;
        elementFilterTxt.disabled=false;
        //elementOUClass.disabled=false;
        var elementTableMC= <HTMLElement> document.getElementById("tblMultiCommit");
        elementTableMC.innerHTML="";
      }
    });

    jQuery("#txtFilter").keyup(function(){
      that.filterFunction();
    });

    //setTimeout(this.ManageDocumentTypeDisplayNameKGS,2000);
    setTimeout(this.BindFunctions, 2000);
  }

  public BindFunctions()
    {
      $("#region option").each(function()
      {
          var strOptVal=$(this).val();
          if( (strOptVal != "defaultOS") && (ObjstoreArr.indexOf(strOptVal) == -1))
          {
            $(this).hide();
          }
      });
  }

  public PopulateObjectStores()
  {
   let objstre: Array<any>;
   objstre = this.properties.ObjectStores.split('\n');
   var option = '';
        for (var i = 0; i < objstre.length; i++) {
          option += '<option value="' + objstre[i] + '">' + objstre[i] + '</option>';
        }
		jQuery('#region').append(option);
   }

  public CreateDefaultTableStructure()
  {
    var tr = document.createElement('tr');
    var fnLabelColn=$(".fnLabel");
    for (var j = 0; j < fnLabelColn.length; j++)
    {
        var th = document.createElement('th');
        var text = document.createTextNode(fnLabelColn[j].innerText);
        th.style.border="3px solid #19d0e3";
        th.appendChild(text);
        tr.appendChild(th);
    }

    var thDelete = document.createElement('th');
    var textDelete = document.createTextNode("Delete");
    thDelete.style.border="3px solid #19d0e3";
    thDelete.appendChild(textDelete);
    tr.appendChild(thDelete);

    tr.style.backgroundColor="#d5dadf";
    var elementTableMC= <HTMLElement> document.getElementById("tblMultiCommit");
    elementTableMC.appendChild(tr);
  }

  public AddRecord()
  {
    var tr = document.createElement('tr');
    var fnInputColn=$(".inputField");
    for (var j = 0; j < fnInputColn.length; j++)
    {
      var elementInput= <HTMLInputElement> document.getElementById(fnInputColn[j].id);
        var td = document.createElement('td');
        //var text = document.createTextNode(elementInput.value);
        var text = elementInput.value;
        if(elementInput.dataset["fieldtype"] =="DateTime")
        {
          if(text){
          //text=moment(text).format("MM-DD-YYYY");
          text=moment(text).format("DD-MMM-YYYY");
        }
        }
        var  inputElemetTxt=document.createElement("INPUT");
        inputElemetTxt.setAttribute("type", "text");
        inputElemetTxt.setAttribute("value", text.toString());
        inputElemetTxt.setAttribute("id", elementInput.id);
        inputElemetTxt.setAttribute("name", elementInput.name);
        inputElemetTxt.setAttribute("alt", elementInput.dataset["fieldtype"]);
        inputElemetTxt.setAttribute("disabled", "true");
       // inputElemetTxt.innerHTML=text.toString();
        td.appendChild(inputElemetTxt);
        td.style.border="2px solid #b14419";
        tr.appendChild(td);
    }
    var tdBtnDel = document.createElement('td');
    var delBtn=document.createElement("IMG");

    delBtn.setAttribute("value", "Delete Row");
    delBtn.setAttribute("onclick","this.parentNode.parentNode.remove();");

    delBtn.setAttribute("src","../SiteAssets/DelButton.JPG");
   // delBtn.setAttribute("src","https://keysighttech.sharepoint.com/sites/Filenet-test/SiteAssets/DelButton.JPG");
    tdBtnDel.appendChild(delBtn);
    tdBtnDel.style.border="2px solid #b14419";
    tr.appendChild(tdBtnDel);

    var elementTableMC= <HTMLElement> document.getElementById("tblMultiCommit");
    elementTableMC.appendChild(tr);
  }

  public InitiateMultiUpload()
  {
    var elementTableMC= <HTMLElement> document.getElementById("tblMultiCommit");
    var mcRows = elementTableMC.getElementsByTagName("tr");
    for(var i=mcRows.length-1;i>0;i--)
    {
        $("#myProgress").show();
        updateProperties={};
        setTimeout(that.uploadFiles(),2000);
    }
    alert("Uploaded Successfully");
    $("#inputFile").val("");
    $("#previewImg").attr("src","");
    $("#previewImg").css("height","1px");
    $("#myProgress").hide();
  }

  public ValidateInputs()
  {
    let reqCheck=0;
        if(jQuery('#region').val()=="defaultOS")
      {
        alert("Please select Object Store");
        reqCheck=1;
      }
      else if(jQuery('#contenttypes').val()=="defaultCT")
      {
        alert("Please select Content Type");
        reqCheck=1;
      }
      else if(jQuery('#terms').val()=="defaultTerms")
      {
        alert("Please select Document Type");
        reqCheck=1;
      }

		  var inpMandatoryCol=$(".requiredInput");
      for(var p=0; p<inpMandatoryCol.length; p++){
        var inpReqID=inpMandatoryCol[p].id;
        if($("#"+inpReqID).val() == "")
        {
          alert("Please fill all the required field");
          reqCheck=1;
          break;
        }
      }

      var flagOptional1="No";
      var inpOptional1Col=$(".optionalInput1");
      for(var pp=0; pp<inpOptional1Col.length; pp++){
        var Opt1ID=inpOptional1Col[pp].id;
        if($("#"+Opt1ID).val() != "")
        {
          flagOptional1="Yes";
        }
      }

      if(flagOptional1=="No")
      {
        var strMSG="";
        var lblOptional1Col=$(".fnLabeloptional1");
        for(var lpp=0; lpp<lblOptional1Col.length; lpp++)
        {
          lblOptional1Col[lpp].id="lblOptional1";
          if(strMSG =="")
          {
            strMSG=$("#lblOptional1")[0].innerText;
          }
          else
          {
            strMSG=strMSG + " OR "+$("#lblOptional1")[0].innerText;
          }
          lblOptional1Col[lpp].id="";
        }

        if(strMSG !="")
        {
          alert("Please fill one of the field from "+ strMSG);
          reqCheck=1;
          //break;
        }
      }

      var flagOptional2="No";
      var inpOptional2Col=$(".optionalInput2");
      for(var qq=0; qq<inpOptional2Col.length; qq++){
        var Opt2ID=inpOptional2Col[qq].id;
        if($("#"+Opt2ID).val() != "")
        {
          flagOptional2="Yes";
        }
      }

      if(flagOptional2=="No")
      {
        var strMSG="";
        var lblOptional2Col=$(".fnLabeloptional2");
        for(var lqq=0; lqq<lblOptional2Col.length; lqq++)
        {
          lblOptional2Col[lqq].id="lblOptional2";
          if(strMSG =="")
          {
            strMSG=$("#lblOptional2")[0].innerText;
          }
          else
          {
            strMSG=strMSG + " OR "+$("#lblOptional2")[0].innerText;
          }
          lblOptional2Col[lqq].id="";
        }

        if(strMSG !="")
        {
          alert("Please fill one of the field from "+ strMSG);
          reqCheck=1;
          //break;
        }
      }

      var flagOptional3="No";
      var inpOptional3Col=$(".optionalInput3");
      for(var rr=0; rr<inpOptional3Col.length; rr++){
        var Opt3ID=inpOptional3Col[rr].id;
        if($("#"+Opt3ID).val() != "")
        {
          flagOptional3="Yes";
        }
      }

      if(flagOptional3=="No")
      {
        var strMSG="";
        var lblOptional3Col=$(".fnLabeloptional3");
        for(var lrr=0; lrr<lblOptional3Col.length; lrr++)
        {
          lblOptional3Col[lrr].id="lblOptional3";
          if(strMSG =="")
          {
            strMSG=$("#lblOptional3")[0].innerText;
          }
          else
          {
            strMSG=strMSG + " OR "+$("#lblOptional3")[0].innerText;
          }
          lblOptional3Col[lrr].id="";
        }

        if(strMSG !="")
        {
          alert("Please fill one of the field from "+ strMSG);
          reqCheck=1;
          //break;
        }
      }
      return reqCheck;
  }

  public getterms()
  {
    var strTermDocTypeID=this.properties.TermDocTypeID;
    if(this.properties.AllDocumentTypes == "Yes")
    {
      this.getTermsetChildren(strTermDocTypeID).then((resp: any) => {
        resp.value.forEach((term: { id: any; }) => {
          this.getTermsetChildren(term.id).then((response) => {
            response.value.forEach((child: { labels: { name: string; }[]; id: string; }) => {
              //jQuery('#terms').append(new Option(child.labels[0].name, child.labels[0].name+"|"+child.id));
              if(DistinctTerms.indexOf(child.labels[0].name)==-1)
              {
                DistinctTerms.push(child.labels[0].name);
                AllTerms.push([child.labels[0].name, child.labels[0].name+"|"+child.id]);
                console.log(child.labels[0].name+"|"+child.id);
              }
            });
            });
        });
      });
    }
    else
    {
      this.getTermsetChildren(strTermDocTypeID).then((response: any) => {
            response.value.forEach((child: { labels: { name: string; }[]; id: string; }) => {
              //jQuery('#terms').append(new Option(child.labels[0].name, child.labels[0].name+"|"+child.id));
              if(DistinctTerms.indexOf(child.labels[0].name)==-1)
              {
                DistinctTerms.push(child.labels[0].name);
                AllTerms.push([child.labels[0].name, child.labels[0].name+"|"+child.id]);
                console.log(child.labels[0].name+"|"+child.id);
              }
            });
      });
    }
  }

  private getTermsetChildren(term: string):Promise<any>
  {
    var strGroupID=this.properties.TermGroupID;
    var strTermsetID=this.properties.TermSetID;
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/v2.1/termStore/groups/"+strGroupID +"/sets/"+strTermsetID+"/terms/"+term+"/children?$orderby=Name",SPHttpClient.configurations.v1)

    .then((response: SPHttpClientResponse) =>
    {
      return response.json();
    });
  }

  private getTermset():Promise<any>
  {
    var strGroupID=this.properties.TermGroupID;
    var strTermsetID=this.properties.TermSetID;
    var strTermDocTypeID=this.properties.TermDocTypeID;
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/v2.1/termStore/groups/"+strGroupID +"/sets/"+strTermsetID+"/terms/"+strTermDocTypeID+"/children?$orderby=Name",SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>
    {
      return response.json();
    });
  }

  public fetchMetadata()
  {
    const container=document.querySelectorAll("#ctfields .inputField");
    for(let c of container as any)
    {
      let fieldval=c.getAttribute("data-fieldtype");
      let req=c.getAttribute("data-required");
      metadataArray.push({internalName:c.name,value:c.value,required:req,fieldtype:fieldval});
    }
    console.log(metadataArray);
    metadataArray.forEach(metadata =>{
      if(metadata.fieldtype=="DateTime")
      {
        updateProperties[metadata.internalName]=moment(metadata.value).toISOString();
      }
      else if(metadata.fieldtype=="Boolean")
      {
        if(metadata.value=="Yes"||metadata.value=="yes"||metadata.value=="True"||metadata.value=="true")
        {
          updateProperties[metadata.internalName]=true;
        }
        else{
          updateProperties[metadata.internalName]=false;
        }

      }
      else if(metadata.fieldtype=="Number")
      {
        updateProperties[metadata.internalName]=parseFloat(metadata.value);
        //updateProperties[metadata.internalName]=100.00;
      }
      else if(metadata.fieldtype=="URL")
      {
        updateProperties[metadata.internalName]={"Url": metadata.value};
      }
      else if(metadata.fieldtype=="TaxonomyFieldType")
        {
          if(metadata.value.split('|').length>1)
          {
            updateProperties[metadata.internalName]=
            {
                __metadata:
                {
                    "type": "SP.Taxonomy.TaxonomyFieldValue" },
                    Label:metadata.value.split('|')[0] ,
                    TermGuid: metadata.value.split('|')[1],
                    WssId: -1
              };
          }
        }
      else
      {
        updateProperties[metadata.internalName]=metadata.value;
      }
    });
    updateProperties['Document_x0020_Type']=termName;
    var xDocType=jQuery('#terms').val().toString();
    updateProperties['TermDocumentType']={ __metadata:
    {
        "type": "SP.Taxonomy.TaxonomyFieldValue" },
         Label:xDocType.split('|')[0] ,
         TermGuid: xDocType.split('|')[1],
         WssId: -1
    };
    updateProperties['ContentTypeId']=ctvalue;
    console.log(updateProperties);
  }

  public uploadFiles()
  {
    try
    {
      var strFileNameCol=uploadFile.name.split('.');
      var strFileName="";
      for(var zz=0;zz<strFileNameCol.length-1;zz++)
      {
        strFileName=strFileName+strFileNameCol[zz];
      }
      var strTimeStamp=moment(Date.now()).format("YYYYMMDDHHmmss").toString();
      var strRandomID=Math.random();
      strFileName=strFileName+strRandomID;
      strFileName = strFileName.replace(/[^a-zA-Z0-9#]/g, '');
      strFileName=strFileName+"_"+strTimeStamp+"."+ strFileNameCol[strFileNameCol.length-1];
      //sp.web.getFolderByServerRelativePath('decodeurl='+"/sites/Filenet-test/"+doclib+).files.add(uploadFile.name, uploadFile, true).then(f => {
      sp.web.getFolderByServerRelativeUrl("/sites/Filenet-test/"+doclib).files.add(strFileName, uploadFile, true).then(f => {
        console.log("File Uploaded");
          //$("#myProgress").hide();
          var strFileURL=f.data.ServerRelativeUrl.toString();
          this.getUploadedFile(strFileURL).then(itemDetails => {
            //var strTest="Test";
            var ItemID=itemDetails.ID;
            this.UpdateUploadedFileItem(doclib,ItemID);
          }).then(res => {
            console.log(res);
          });
      });
    }
    catch (error)
    {
      $("#myProgress").hide();
      alert("Issue in Uploading the file. Request you to check with SharePoint Administrator.");
    }
  }

  public getUploadedFile(strFileURL:any):Promise<any>
  {
    try
    {
      strFileURL=strFileURL.replace(/#/g,"%23");
      strFileURL=strFileURL.replace(/&/g,"%26");
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/GetFileByServerRelativePath(decodedurl='"+strFileURL+"')/ListItemAllFields",SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) =>
      {
        return response.json();
      });
    } catch (error)
    {
      $("#myProgress").hide();
      alert("Issue in fetching the uploaded file. Request you to check with SharePoint Administrator.");
    }
  }

 public async UpdateUploadedFileItem(doclib:any,strItemID:any)
  {
   try
   {
      var elementMultiCommitChk = <HTMLInputElement> document.getElementById("chkMulticommit");

      if(elementMultiCommitChk.checked == true)
      {
        var elementTableMC= <HTMLElement> document.getElementById("tblMultiCommit");
        var mcRows = elementTableMC.getElementsByTagName("tr");
        updateProperties={};
        var mcRowCells=mcRows[mcRows.length-1].getElementsByTagName("td");
        for(var j=0; j<mcRowCells.length-1;j++)
        {
          var inpElementRef=mcRowCells[j].getElementsByTagName("input")[0];
          if(inpElementRef.value.length>0)
          {
            if(inpElementRef["alt"]=="DateTime")
            {
              updateProperties[inpElementRef.name]=moment(inpElementRef.value).toISOString();
            }
            else if(inpElementRef["alt"]=="Boolean")
            {
              if(inpElementRef.value=="Yes"||inpElementRef.value=="yes"||inpElementRef.value=="True"||inpElementRef.value=="true")
              {
                updateProperties[inpElementRef.name]=true;
              }
              else
              {
                updateProperties[inpElementRef.name]=false;
              }
            }
            else if(inpElementRef["alt"]=="Number")
            {
              updateProperties[inpElementRef.name]=parseFloat(inpElementRef.value);
              //updateProperties[inpElementRef.name]=100.00;
            }
            else if(inpElementRef["alt"]=="URL")
            {
              updateProperties[inpElementRef.name]={"Url": inpElementRef.value};
            }
            else if(inpElementRef["alt"]=="TaxonomyFieldType")
            {
              if(inpElementRef.value.split('|').length>1)
              {
                updateProperties[inpElementRef.name]=
                {
                    __metadata:
                    {
                      "type": "SP.Taxonomy.TaxonomyFieldValue" },
                      Label:inpElementRef.value.split('|')[0] ,
                      TermGuid: inpElementRef.value.split('|')[1],
                      WssId: -1
                };
              }
            }
            else
            {
              updateProperties[inpElementRef.name]=inpElementRef.value;
            }
          }
        }
        updateProperties['Document_x0020_Type']=termName;
        var xDocType=jQuery('#terms').val().toString();
        updateProperties['TermDocumentType']=
        {
          __metadata:
          {
              "type": "SP.Taxonomy.TaxonomyFieldValue" },
              Label:xDocType.split('|')[0] ,
              TermGuid: xDocType.split('|')[1],
              WssId: -1
        };
        updateProperties['ContentTypeId']=ctvalue;
        mcRows[mcRows.length-1].remove();
        console.log(updateProperties);
      }
      var siteUrl = this.context.pageContext.web.absoluteUrl ;
      let web = Web(siteUrl);
       await web.lists.getByTitle(doclib).items.getById(strItemID).update(updateProperties ).then(i => {
        console.log(i);
      });
      $("#myProgress").hide();

      //var elementMultiCommitChk = <HTMLInputElement> document.getElementById("chkMulticommit");
      if(elementMultiCommitChk.checked == false)
      {
        alert("Uploaded Successfully");
        $("#inputFile").val("");
        $("#previewImg").attr("src","");
        $("#previewImg").css("height","1px");
      }
   } catch (error)
   {
    $("#myProgress").hide();
    alert("Issue in updating the file Properties. Request you to check with SharePoint Administrator.");
   }

  }

  public callContentTypeAPI(selvalue: any)
  {
    //alert(selvalue);
    jQuery('#contenttypes').empty();
    jQuery('#contenttypes').append(new Option("--Select--", "defaultCT"));
    let option:any;
    this.getContentTypes(selvalue).then((response) => {
      console.log(response);
      response.value.forEach((opt: { Name: string; Id: { StringValue: string; }; }) => {
        jQuery('#contenttypes').append(new Option(opt.Name, opt.Id.StringValue));
      });

    });
  }

  public callContentTypeFieldsAPI(ctvalue: any,listvalue: any)
  {
    jQuery('#ctfields').empty();
    //let propertyPaneRestrictedFields="";

    let Restrictedproparray=(this.properties.RestrictedFields).split('\n');
    this.getContentTypesFields(ctvalue,listvalue).then((response) => {
      console.log(response);
      let html:any=``;
      let count=0;
      // Get Array of Required fields
      response.value.forEach((field: { Title: any; InternalName: any; Required:any; TypeAsString:any; }) => {
        var strTitle=field.Title;

        if(strTitle == "SO_Number")
        {
          strTitle="Keysight Order Number";
        }
        if(Restrictedproparray.indexOf(field.Title)==-1)
        {
          if(count%3==0)
          {
            html+=`<div class="row" style="padding-top: 1%;">`;
          }
          if(field.InternalName!="Term_x0020_Document_x0020_Type"&&field.InternalName!="TermDocumentType"&&field.InternalName!="Document_x0020_Type"&&field.InternalName!="Document_Type")
          {

            if(strTitle== "OU Name (Doc Class)")
            {
              html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${strTitle}</label></div>
              <div class="col-sm-6 col-md-6 col-lg-6"><select id="OUClassName" style="width: 158px;" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}" class="inputField" name="${field.InternalName}"></select></div></div></div>`;
            }
            else if(field.Required )
            {
              if(field.TypeAsString=="DateTime")
              {
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${strTitle} <span style="color:red"><b>*</b></span></label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><input type="date" id="${field.InternalName}" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}" class="inputField" name="${field.InternalName}" ></div></div></div>`;
              }
              else if(field.TypeAsString=="Boolean")
              {
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${strTitle} <span style="color:red"><b>*</b></span></label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><select id="${field.InternalName}" style="width: 158px;" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}" class="inputField" name="${field.InternalName}"><option value="Yes">Yes</option><option value="No">No</option></select></div></div></div>`;
              }
              else if(field.TypeAsString=="TaxonomyFieldType")
              {
                let options:string="";
                let fieldName:string="#"+field.InternalName;
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${field.Title} <span style="color:red"><b>*</b></span></label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><select id="${field.InternalName}" style="width: 158px;" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}"  class="inputField" name="${field.InternalName}"><option value="defaultOption">--Select--</option>${options}</select></div></div></div>`;
                this.getTermsData(field.InternalName).then((response) => {
                  console.log(response);
                  response.value.forEach((value) => {
                    //managedFieldValue.push({Title:value.labels[0].name,value:value.id});

                    jQuery('#'+field.InternalName).append(new Option(value.labels[0].name,value.labels[0].name+"|"+ value.id));
                  });
                });
              }
              else
              {
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${strTitle} <span style="color:red"><b>*</b></span></label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><input type="text" id="${field.InternalName}" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}" class="inputField" name="${field.InternalName}"></div></div></div>`;
              }
           }
           else
           {
            if(field.TypeAsString=="DateTime")
              {
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${strTitle}</label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><input type="date" id="${field.InternalName}" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}" class="inputField" name="${field.InternalName}" ></div></div></div>`;
              }
              else if(field.TypeAsString=="Boolean")
              {
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${strTitle}</label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><select id="${field.InternalName}" style="width: 158px;" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}" class="inputField" name="${field.InternalName}"><option value="Yes">Yes</option><option value="No">No</option></select></div></div></div>`;
              }
              else if(field.TypeAsString=="TaxonomyFieldType")
              {
                let options:string="";
                let fieldName:string="#"+field.InternalName;
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${field.Title}</label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><select id="${field.InternalName}" style="width: 158px;" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}"  class="inputField" name="${field.InternalName}"><option value="defaultOption">--Select--</option>${options}</select></div></div></div>`;
                this.getTermsData(field.InternalName).then((response) => {
                  console.log(response);
                  response.value.forEach((value) => {
                    //managedFieldValue.push({Title:value.labels[0].name,value:value.id});

                    jQuery('#'+field.InternalName).append(new Option(value.labels[0].name, value.labels[0].name+"|"+ value.id));
                  });
                });
              }
              else
              {
                html+=`<div class="col-sm-4 col-md-4 col-lg-4"><div class="row"><div class="col-sm-6 col-md-6 col-lg-6"><label class="fnLabel" for="${strTitle}" style="font-weight: 400">${strTitle}</label></div>
                <div class="col-sm-6 col-md-6 col-lg-6"><input type="text" id="${field.InternalName}" data-fieldtype="${field.TypeAsString}" data-required="${field.Required}" class="inputField" name="${field.InternalName}"></div></div></div>`;
              }
            }
          }
          else
          {
            count--;
          }
          if(count%3==2)
          {
            html+=`</div>`;
          }
          count++;
        }
      });
      jQuery('#ctfields').append(html);
      var d = new Date();
     // var currDAte=moment(Date.now()).format("DD-M-YYYY").toString();
      if($('#Date_x0020_Filed').length >0)
      {
        $('#Date_x0020_Filed').val(d.toISOString().substring(0, 10));
       // $('#Date_x0020_Filed').val(currDAte);
        var elementDateFiled = <HTMLInputElement> document.getElementById("Date_x0020_Filed");
        elementDateFiled.disabled = true;
      }

      if($('#Document_x0020_Date').length >0)
      {
        $('#Document_x0020_Date').val(d.toISOString().substring(0, 10));
      }
      this.PopulateOUClasses();
    });
  }

  private getContentTypes(selvalue: string):Promise<any>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('"+selvalue+"')/contenttypes?$orderby=Name",SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>
    {
      return response.json();
    });
  }

  private getContentTypesFields(ctvalue: string,listvalue: string):Promise<any>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('"+listvalue+"')/contenttypes('"+ctvalue+"')/fields?$filter=Group eq 'Filenet'",SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>
    {
      return response.json();
    });
  }

  public  GetObjectLists():Promise<any>
  {
    try
    {
      debugger;
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl+"/_api/web/lists?$select=Title,IsApplicationList&$expand=properties&$filter=IsApplicationList eq false&$top=4995", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
      //async: true;
      return response.json();
      });
    }
    catch
    {
    }
  }

  public async GetAllConfigFields()
  {
    try
    {
      var siteUrl = this.context.pageContext.web.absoluteUrl ;
      let web = Web(siteUrl);
      const items: any[] = await web.lists.getByTitle("ADDOCFieldsConfigList").items.select("Title", "field_AssociatedContentType","field_DocumentType","field_FieldCollection1","field_FieldCollection2","field_FieldCollection3").orderBy("Title").getAll();
      for(var x=0;x<items.length;x++)
      {
        AllConfigFields.push([items[x].Title.trim(),items[x].field_AssociatedContentType,items[x].field_DocumentType,items[x].field_FieldCollection1,items[x].field_FieldCollection2,items[x].field_FieldCollection3]);

        if(AllDoumentTypes.indexOf(items[x].field_DocumentType)==-1)
        {
         AllDoumentTypes.push(items[x].field_DocumentType);
        }

        CTDocTypes.push([items[x].Title,items[x].field_DocumentType]);
      }
      console.log(items);
    }
    catch (error)
    {
       // alert(error);
    }
  }

  public GetObjectStoreConfigFields(selecteObjStore:string)
  {
    try
    {
      debugger;
      ObjstoreConfigFields=[];
      for(var j=0;j<AllConfigFields.length; j++)
      {
        if(AllConfigFields[j][0] == selecteObjStore)
        {
          ObjstoreConfigFields.push([AllConfigFields[j][1].trim(),AllConfigFields[j][2],AllConfigFields[j][3],AllConfigFields[j][4],AllConfigFields[j][5],AllConfigFields[j][6]]);
        }
      }
    }
    catch
    {
    }
  }

  public GetCTConfigFields(selecteDCT:string)
  {
    try
    {
      debugger;
      CTConfigFields=[];
      for(var j=0;j<ObjstoreConfigFields.length; j++)
      {
        if((ObjstoreConfigFields[j][0] == selecteDCT) || (ObjstoreConfigFields[j][0] == "NotApplicable"))
        {
          CTConfigFields.push([ObjstoreConfigFields[j][1].trim(),ObjstoreConfigFields[j][2],ObjstoreConfigFields[j][3],ObjstoreConfigFields[j][4],ObjstoreConfigFields[j][5]]);
        }
      }
    }
    catch
    {
    }
  }

  public GetDocumentTypeFields(selectedDocType:string)
  {
    try
    {
      debugger;
      DocumentTypeConfigFields=[];
      for(var j=0;j<CTConfigFields.length; j++)
      {
        if((CTConfigFields[j][0] == selectedDocType) || (CTConfigFields[j][0] == "NotApplicable"))
        {
          var FieldCollection1=CTConfigFields[j][1];
          var FieldCollection2=CTConfigFields[j][2];
          var FieldCollection3=CTConfigFields[j][3];

          var FieldCollection1Arr=[];
          var FieldCollection2Arr=[];
          var FieldCollection3Arr=[];

          if(FieldCollection1)
          {
              if(FieldCollection1.toString().indexOf("&&")>0)
              {
                FieldCollection1Arr=(FieldCollection1.toString()).split("&&");
                for(var x=0;x<FieldCollection1Arr.length;x++)
                {
                  DocumentTypeConfigFields.push([FieldCollection1Arr[x],"requiredInput"]);
                }
              }
              else if(FieldCollection1.toString().indexOf("||")>0)
              {
                FieldCollection1Arr=(FieldCollection1.toString()).split("||");
                for(var x=0;x<FieldCollection1Arr.length;x++)
                {
                  DocumentTypeConfigFields.push([FieldCollection1Arr[x],"optionalInput1"]);
                }
              }
              else
              {
                DocumentTypeConfigFields.push([FieldCollection1,"requiredInput"]);
              }
          }

          if(FieldCollection2)
          {
              if(FieldCollection2.toString().indexOf("&&")>0)
              {
                FieldCollection2Arr=(FieldCollection2.toString()).split("&&");
                for(var x=0;x<FieldCollection2Arr.length;x++)
                {
                  DocumentTypeConfigFields.push([FieldCollection2Arr[x],"requiredInput"]);
                }
              }
              else if(FieldCollection2.toString().indexOf("||")>0)
              {
                FieldCollection2Arr=(FieldCollection2.toString()).split("||");
                for(var x=0;x<FieldCollection2Arr.length;x++)
                {
                  DocumentTypeConfigFields.push([FieldCollection2Arr[x],"optionalInput2"]);
                }
              }
              else
              {
                DocumentTypeConfigFields.push([FieldCollection2,"requiredInput"]);
              }
          }

          if(FieldCollection3)
          {
              if(FieldCollection3.toString().indexOf("&&")>0)
              {
                FieldCollection3Arr=(FieldCollection3.toString()).split("&&");
                for(var x=0;x<FieldCollection3Arr.length;x++)
                {
                  DocumentTypeConfigFields.push([FieldCollection3Arr[x],"requiredInput"]);
                }
              }
              else if(FieldCollection3.toString().indexOf("||")>0)
              {
                FieldCollection3Arr=(FieldCollection3.toString()).split("||");
                for(var x=0;x<FieldCollection3Arr.length;x++)
                {
                  DocumentTypeConfigFields.push([FieldCollection3Arr[x],"optionalInput3"]);
                }
              }
              else
              {
                DocumentTypeConfigFields.push([FieldCollection3,"requiredInput"]);
              }
          }

        }
      }
    }
    catch(error)
    {
     // alert(error);
    }
  }

  public ManageDocumentTypeFields()
  {
    var spnMandatoryHTML= "";
    try
    {
      debugger;
      // Delete all mandatory Spans if already added
      var spnMandatoryCol=$("span.spnMandatory");
      for(var s=0; s<spnMandatoryCol.length; s++){
        spnMandatoryCol[s].id="spnTempTBD";
        $("#spnTempTBD").remove();
      }

      var inpMandatoryCol=$(".requiredInput");
      for(var p=0; p<inpMandatoryCol.length; p++){
        var inpReqID=inpMandatoryCol[p].id;
        $("#"+inpReqID).removeClass("requiredInput");
      }

      var inpOptionalCol1=$(".optionalInput1");
      for(var a=0; a<inpOptionalCol1.length; a++){
        var inpOptnl1ID=inpOptionalCol1[a].id;
        $("#"+inpOptnl1ID).removeClass("optionalInput1");
      }

      var inpOptionalCol2=$(".optionalInput2");
      for(var b=0; b<inpOptionalCol2.length; b++){
        var inpOptnl2ID=inpOptionalCol2[b].id;
        $("#"+inpOptnl2ID).removeClass("optionalInput2");
      }

      var inpOptionalCol3=$(".optionalInput3");
      for(var c=0; c<inpOptionalCol3.length; c++){
        var inpOptnl3ID=inpOptionalCol3[c].id;
        $("#"+inpOptnl3ID).removeClass("optionalInput3");
      }

      // Remove LAbel styles - Required
      var lblMandatoryCol=$(".fnLabelBold");
      for(var q=0; q<lblMandatoryCol.length; q++){
        lblMandatoryCol[q].id="lblFnID";
        $("#lblFnID").css("font-weight","400");
        $("#lblFnID").removeClass("fnLabelBold");
        lblMandatoryCol[q].id="";
      }

	  // Remove LAbel styles - Optional
	  var lblOptional1Col=$(".fnLabeloptional1");
      for(var d=0; d<lblOptional1Col.length; d++){
        lblOptional1Col[d].id="lblFnID";
        $("#lblFnID").css("color","#121212");
        $("#lblFnID").removeClass("fnLabeloptional1");
        lblOptional1Col[d].id="";
      }

	  var lblOptional2Col=$(".fnLabeloptional2");
      for(var e=0; e<lblOptional2Col.length; e++){
        lblOptional2Col[e].id="lblFnID";
        $("#lblFnID").css("color","#121212");
        $("#lblFnID").removeClass("fnLabeloptional2");
        lblOptional2Col[e].id="";
      }


	  var lblOptional3Col=$(".fnLabeloptional3");
      for(var f=0; f<lblOptional3Col.length; f++){
        lblOptional3Col[f].id="lblFnID";
        $("#lblFnID").css("color","#121212");
        $("#lblFnID").removeClass("fnLabeloptional3");
        lblOptional3Col[f].id="";
      }


      // Delete all mandatory class from input fields if already added
      //CTRequiredFields=[];
      for(var j=0;j<DocumentTypeConfigFields.length; j++)
      {
        // add mandatory span
       var strTempControlIdentifier=DocumentTypeConfigFields[j][0].trim();
       var strTempLinkCol=$("label.fnLabel:contains('"+strTempControlIdentifier+"')");
       if(strTempLinkCol.length > 0)
      {
		    var strClass=DocumentTypeConfigFields[j][1];
        for(var sr=0; sr<strTempLinkCol.length; sr++ )
        {
          var strTemplnkControl=strTempLinkCol[sr];
          var strTempControlText=strTemplnkControl.innerText;
          if( (typeof strTemplnkControl !== "undefined") && (strTempControlText == strTempControlIdentifier ))
          {
            //strTemplnkControl.id="spnTemp";
            var ctrlInputID=strTemplnkControl.parentElement.nextElementSibling.firstElementChild.id;
            if(strClass == "requiredInput")
            {
              strTemplnkControl.id="fnLablelID";
              $('#fnLablelID').addClass('fnLabelBold');
              $('#fnLablelID').css("font-weight","700");
              strTemplnkControl.id="";
              $("<span  class='spnMandatory' style='color:red'><b>*</b></span>", {html: ""}).insertAfter(strTemplnkControl);

              $('#'+ctrlInputID).addClass('requiredInput');
            }
            else if(strClass == "optionalInput1")
            {
              strTemplnkControl.id="fnLablelID";
              $('#fnLablelID').addClass('fnLabeloptional1');
              $('#fnLablelID').css("color","#20c715");
              strTemplnkControl.id="";
              $('#'+ctrlInputID).addClass('optionalInput1');
            }
            else if(strClass == "optionalInput2")
            {
              strTemplnkControl.id="fnLablelID";
              $('#fnLablelID').addClass('fnLabeloptional2');
              $('#fnLablelID').css("color","#15c77d");
              strTemplnkControl.id="";
              $('#'+ctrlInputID).addClass('optionalInput2');
            }
            else if(strClass == "optionalInput3")
            {
              strTemplnkControl.id="fnLablelID";
              $('#fnLablelID').addClass('fnLabeloptional3');
              $('#fnLablelID').css("color","#FF9800");
              strTemplnkControl.id="";
              $('#'+ctrlInputID).addClass('optionalInput3');
            }
          }
        }
      }
      }
    }
    catch
    {
    }
  }

  public async GetAllOUClasses()
  {
    try
    {
        var siteUrl = this.context.pageContext.web.absoluteUrl ;
        let web = Web(siteUrl);
        const items: any[] = await web.lists.getByTitle("ADDOCOUClasses").items.select("Title", "OUClassName").orderBy("OUClassName").get();
        for(var x=0;x<items.length;x++)
        {
          AllOUClasses.push([items[x].Title,items[x].OUClassName]);
        }
        console.log(items);
    }
    catch
    {
    }
  }

  public GetObjectStoreOUClasses(selectedCobjStore:string)
  {
    try
    {
      debugger;
      RegionalOUClasses=[];
      for(var j=0;j<AllOUClasses.length; j++)
      {
        if(AllOUClasses[j][0] == selectedCobjStore)
        {
          RegionalOUClasses.push(AllOUClasses[j][1]);
        }
      }
    }
    catch
    {
    }
  }

  public PopulateOUClasses()
  {
    //OUClassName
    try
    {
      RegionalOUClasses.sort();
      $("#OUClassName").empty();
      for(var k=0;k<RegionalOUClasses.length; k++)
      {
        $('#OUClassName').append("<option>" + RegionalOUClasses[k].trim() + "</option>");
      }
    }
    catch
    {
    }
  }

  public ManageDocumentTypeDisplayNameKGS()
  {
    if(AllDoumentTypes.length>0)
      {
        var DocTypesAvailable=$("#terms option");
        if(DocTypesAvailable.length>1)
        {
          for(var zz=1; zz<DocTypesAvailable.length; zz++)
          {
            var optDisplay=$("#terms option")[zz].innerText;
            if(AllDoumentTypes.indexOf(optDisplay) !=-1)
            {
              $("#terms option")[zz].innerText="KGS-"+optDisplay;
            }
          }
        }
      }
      $("#txtFilter").val("KGS");
      that.filterFunction();
  }

  public filterFunction()
  {
    var input, filter, ul, li, a, i;
    filter = $('#txtFilter').val().toString().toUpperCase();
    //let selvalue = jQuery('#region').val();

   // if((selvalue == "FNAMERICA") || (selvalue == "FNASIA") || (selvalue == "FNEUROPE"))
    //{
      //if(filter=="")
      //{
        //filter="KGS-";
      //}
    //}
    //filter = input.value.toUpperCase();
    var div = $('#terms');
    //a = div.getElementsByTagName("option");
    a = div.find("option");
    for (i = 0; i < a.length; i++) {
      var txtValue = a[i].textContent || a[i].innerText;
      if (txtValue.toUpperCase().indexOf(filter) > -1) {
        a[i].style.display = "";
      } else {
        a[i].style.display = "none";
      }
    }
  }


  public ManageCTDocTypes(selectedCobjStore:string) //,selectedDC:string
  {
      try
      {
        var CTDOcsAvailable=[];
        if((selectedCobjStore != "FNASIA") && (selectedCobjStore != "FNAMERICA") && (selectedCobjStore != "FNEUROPE"))
        {
          $("#terms option").show();
          //alert("Non-KGS Type");
        }
        else
        {

          $("#terms option").hide();
          $("#terms option").addClass("tbdOptions");
          var termColnKGS= $("#terms option");
          //alert("KGS Type");
          for(var u=0;u<CTDocTypes.length;u++)
          {
              if(CTDocTypes[u][0] == selectedCobjStore)
              {
                CTDOcsAvailable.push(CTDocTypes[u][1]);
              }
          }

          if(CTDOcsAvailable.length>0)
          {
            termColnKGS[0].style.display = "block";
            termColnKGS[0].id="tmpOptID";
            $("#tmpOptID").removeClass("tbdOptions");
            termColnKGS[0].id="";

            for(var v=0;v<CTDOcsAvailable.length;v++)
            {
              for(var t=1;t<termColnKGS.length;t++)
              {
                var strOptText=termColnKGS[t].innerText;
                //var strCTDocAvlText="KGS-"+CTDOcsAvailable[v];
                var strCTDocAvlText=CTDOcsAvailable[v];
                if((strOptText== strCTDocAvlText))
                {
                  termColnKGS[t].style.display = "block";
                  termColnKGS[t].id="tmpOptID";
                  $("#tmpOptID").removeClass("tbdOptions");
                  termColnKGS[t].id="";
                }
              }
            }
          }

          termColnKGS= $(".tbdOptions");
          termColnKGS.remove();

        }
      } catch (error)
      {

      }
  }

  public getTermsData(termset:string):Promise<any>
  {
      //  Get the Term ID from Property Pane
      var TermTempID="";
      let TermColn : Array<any>;
      TermColn = this.properties.TermCollection.split('\n');
      if(TermColn.length>0)
      {
        for (var i=0; i<TermColn.length; i++)
        {
           var strTermSetTemp=TermColn[i].split('|')[0].trim();
           if(strTermSetTemp == termset)
           {
              TermTempID =TermColn[i].split('|')[1].trim();
           }
        }
      }

      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl+"/_api/v2.1/termStore/groups/"+ this.properties.TermGroupID+"/sets/"+ this.properties.TermSetID+"/terms/"+TermTempID+"/children",SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) =>
      {
        return response.json();
      });
  }

  protected getdataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getdisableReactivePropertyChanges(): boolean {
    return true;
  }

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
                }),

                PropertyPaneTextField('TermGroupID', {
                  label:"TermGroupID"
                }),

                PropertyPaneTextField('AllDocumentTypes', {
                  label: "AllDocumentTypes"
                }),

                PropertyPaneTextField('TermSetID', {
                  label: "TermSetID"
                }),

                PropertyPaneTextField('TermDocTypeID', {
                  label: "TermDocTypeID"
                }),


                PropertyPaneTextField('UnsupportedPreviewFileTypes', {
                  label: "Unsupported Preview File Types",
                  multiline: true
                }),

                PropertyPaneTextField('TermCollection', {
                  label:"TermCollection",
                  multiline: true
                }),

                PropertyPaneTextField('ObjectStores', {
                  label: "ObjectStores",
                  multiline: true
                }),

                PropertyPaneTextField('RestrictedFields', {
                  label: "Restricted Fields",
                  multiline: true
                })
              ]
            }
          ]
        }
      ]
    };
  }

}
