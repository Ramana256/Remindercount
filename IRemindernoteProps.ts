import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';

export interface IRemindernoteProps {
  spHttpClient: SPHttpClient;
  Title: string;
  TimetoRemind: string;
  ProjectsArray: Array<string>[];
  siteurl: string;
  RepeatInterval: string;
  UploadedFilesArray: Array<string>[];
 
 
}

