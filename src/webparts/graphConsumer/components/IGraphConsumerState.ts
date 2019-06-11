import { IUserItem } from './IUserItem';
import { ILibrary } from './ILibrary';
import { IDocument } from './IDocument';

export interface IGraphConsumerState {
    users: Array<IUserItem>;
    searchFor: string;
    libraries:Array<ILibrary>;
    documents:Array<IDocument>;
    selectionDetails:any;
    loadComplete:boolean;
    errorDialog:boolean;
    errorMessage:string;
  }