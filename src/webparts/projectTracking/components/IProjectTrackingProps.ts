import { WebPartContext } from '@microsoft/sp-webpart-base';  

import {
  ButtonClickedCallback,mynewnumber,MYchoices
} from '../../../models';
export interface IProjectTrackingProps {
  description: string;
  context: WebPartContext; 
  onAddButton?: ButtonClickedCallback;
  onDeleteBtn?: ButtonClickedCallback;
  onmynumber?: mynewnumber[];
  mychoices?:MYchoices[];
  mychoices2?:MYchoices[];
  mychoices3?:MYchoices[];
  mychoices4?:MYchoices[];
}
export interface ISpFxHttpClientDemoProps {
  onmynumber?: mynewnumber[];
  onAddButton?: ButtonClickedCallback;
  onDeleteBtn?: ButtonClickedCallback; 
}
