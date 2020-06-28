import styles from './Remindernote.module.scss';
import { IRemindernoteProps } from './IRemindernoteProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as React from 'react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import Editor from 'react-medium-editor';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { ISPListData } from '@microsoft/sp-page-context/lib/SPList';
import * as Datetime from 'react-datetime';
import 'react-datetime/css/react-datetime.css';


import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import {
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';
import {
  SPHttpClient,
  SPHttpClientBatch,
  SPHttpClientResponse, SPHttpClientConfiguration
} from '@microsoft/sp-http';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';


import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { Promise } from 'es6-promise';
import * as lodash from 'lodash';
import * as jquery from 'jquery';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import ReactFileReader from 'react-file-reader';

import { Checkbox, CheckboxGroup } from 'react-checkbox-group';

import { default as pnp, ItemAddResult, Web } from "sp-pnp-js";

require('medium-editor/dist/css/medium-editor.css');
require('medium-editor/dist/css/themes/default.css');


var App = React.createClass({

  getInitialState() {
    return { text: 'Enter Rich Text Description' };
  },

  render() {
    var divStyle = {
      background: "#eee",
      padding: "10px",
      margin: "1px",
      width: "100%",
      height: "140px",
    };

    return (
      <div style={divStyle}>
        <Editor
          text={this.state.text}
          onChange={this.props.handleChange}
        />
      </div>
    );
  },
  handleChange(text, medium) {
    this.setState({ text: text });
  }
});


export default class Remindernote extends React.Component<IRemindernoteProps, {}> {

  public state: IRemindernoteProps;
  constructor(props, context) {
    super(props);
    this.state = {
      spHttpClient: this.props.spHttpClient,
      description: "",
      siteurl: this.props.siteurl,
      Title: "",
      RepeatInterval: "",      
      TimetoRemind: "",
      ItemGuid: this.GenerateGuid().toString(),
      loading: false,
      UploadedFilesArray: [],
      ProjectsArray: [],
  
    };
    this.onChangeDeleteDocument = this.onChangeDeleteDocument.bind(this);
  }

  private CreateNewItem(): void {
  }
  onChangeDeleteDocument(val) {
    var array = this.state.UploadedFilesArray;
    var MainIndex = val.currentTarget.dataset.id.toString(); // Let's say it's Bob.
    var indextoDelete = 0;

    for (var i = 0; i < array.length; i++) {
      if (array[i] != undefined) {
        var temp = array[i].toString().split('|');
        if (temp[1] == MainIndex) {
          indextoDelete = i;

        }
      }
    }
    delete array[indextoDelete];
    this.state.UploadedFilesArray = [];
    this.state.UploadedFilesArray = array;
    this.setState({ UploadedFilesArray: array });
    let getData = [];
    let str = [];
    for (let i = 0; i < this.state.UploadedFilesArray.length; i++) {
      if (this.state.UploadedFilesArray[i] != undefined) {
        var tempx = this.state.UploadedFilesArray[i].toString().split('|');
        str.push(<li key={tempx[0]} onClick={this.onChangeDeleteDocument.bind(this)} data-id={tempx[1]}> Uploaded File : {tempx[0]} - <a className={styles.MyHeadingsAnchor}>Delete </a></li>);
      }
    }
    getData.push(<ul>{str}</ul>);
    this.updateDocumentLibrary(MainIndex);


  }


  private updateDocumentLibrary(MainIndex) {

    var reactHandler = this;
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("SitePages", "");
    console.log(NewSiteUrl);

    const body: string = JSON.stringify({
      '__metadata': {
        'type': `SP.Data.MyDocsItem`
      },
      'Deleted': `Yes`
    });
    return this.props.spHttpClient.post(`${NewSiteUrl}/_api/web/lists/getbytitle('MyDocs')/items(${MainIndex})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': "*",
          'X-HTTP-Method': 'MERGE'
        },
        body: body
      }).then((response) => {
        console.log(response);
      });


  }

  GenerateGuid() {
    var date = new Date();
    var guid = date.valueOf();
    return guid;
  }

  componentDidMount() {
   this._renderListAsync();

  }

  private _renderListAsync(): void {
    var reactHandler = this;
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("SitePages", "");
    console.log(NewSiteUrl);
    jquery.ajax({
      url: `${NewSiteUrl}/_api/web/lists/getbytitle('Reminder Notes')/items?&$select=Title,ID`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        var myObject = JSON.stringify(resultData.d.results);
        this.setState({ ProjectsArray: resultData.d.results,
                        Disciplined:"Civil",
          })
      }.bind(this),
      error: function (jqXHR, textStatus, errorThrown) {
      }

    });
  }






  handleFiles = files => {
    var TemFileGuidName = [];
    var component = this;
    component.setState({ loading: true });
    var FileExtension = this.getFileExtension1(files.fileList[0].name);
    var date = new Date();
    var guid = date.valueOf();
    if (this.state.ItemGuid == "-1") {
      this.setState({ ItemGuid: guid });
    }
    //alert(this.state.ItemGuid);   
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    console.log(NewSiteUrl);
    let webx = new Web(NewSiteUrl);

    var FinalName = guid + FileExtension;


    webx.get().then(r => {
      var myBlob = this._base64ToArrayBuffer(files.base64);
      webx.getFolderByServerRelativeUrl("MyDocs")
        .files.add(FinalName.toString(), myBlob, true)
        .then(function (data) {
          var RelativeUrls = "MyDocs/" + FinalName;//files.fileList[0].name;
          webx.getFolderByServerRelativeUrl(RelativeUrls).getItem().then(item => {
            // updating Start
            TemFileGuidName[0] = files.fileList[0].name + "|" + item["ID"];
            webx.lists.getByTitle("MyDocs").items.getById(item["ID"]).update({
              Guid0: guid.toString(),
              ActualName: files.fileList[0].name
            }).then(r => {
              component.setState({ loading: false });
              component.setState({ UploadedFilesArray: component.state.UploadedFilesArray.concat(TemFileGuidName) });
            });
          }); //Retrive Doc Info End
        });
    });
  }

  private getFileExtension1(filename) {
    return (/[.]/.exec(filename)) ? /[^.]+$/.exec(filename)[0] : undefined;
  }


  private _base64ToArrayBuffer(base64) {
    var binary_string = window.atob(base64.split(',')[1]);
    var len = binary_string.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) {
      bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
  }


  public onSelectTimetoremind(event: any): void {
    this.setState({ TimetoRemind: event._d });
  }
  public OnchangeTitle(event: any): void {
    this.setState({ Title: event.target.value });
  }

  public OnchangeRepeatInterval(event: any): void {
    this.setState({ RepeatInterval: event.target.value });
  }

  public render(): React.ReactElement<IRemindernoteProps> {
    let content;
    if (this.state.loading) {
      content = <div><img src="https://*****************/sites/dev/SiteAssets/loadingnew.gif" /></div>;
    } else {
      // content = <div>{this.state.UploadedFilesArray}</div>;
    }


    let getData = [];
    let str = [];
    for (let i = 0; i < this.state.UploadedFilesArray.length; i++) {
      if (this.state.UploadedFilesArray[i] != undefined) {
        var tempx = this.state.UploadedFilesArray[i].toString().split('|');
        str.push(<li key={tempx[0]} onClick={this.onChangeDeleteDocument.bind(this)} data-id={tempx[1]}> Uploaded File : {tempx[0]} - <a className={styles.MyHeadingsAnchor}>Delete </a></li>);
      }
    }
    getData.push(<ul>{str}</ul>);


   
    return (
      <div className={styles.addNewRfi} >
        <h1> SPFX - Complex Form </h1>
        <div>
          {content}
          {getData}
        </div>
      
        <div className={styles.row}>
          <div className={styles.label}>
            Input Value 1
           </div>
          <div >
            <input type="text" className={styles.myinput} value={this.state.Title} onChange={this.OnchangeTitle.bind(this)} />
          </div>
        </div>

        <div className={styles.rowDate}>
          <div className={styles.label}>
            Date
           </div>
          <div className={styles.myinput}>
            <Datetime onChange={this.onSelectTimetoremind.bind(this)} />
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.label}>
            2)Input Value 2
           </div>
          <div >
            <input type="text" className={styles.myinput} value={this.state.RepeatInterval} onChange={this.OnchangeRepeatInterval.bind(this)} />
          </div>
        </div>

        <div className={styles.row}>
          <ReactFileReader fileTypes={[".csv", ".xlsx", ".Docx"]} handleFiles={this.handleFiles.bind(this)} base64={true} >
            <button className='btn'>Upload</button>
          </ReactFileReader>


        </div>



        <div className={styles.row}>
          <div  >
            <button id="btn_add" className={styles.button} onClick={this.CreateNewItem.bind(this)}>Create New ListItem </button>
          </div>
        </div>

      </div >
    );
  }
}
