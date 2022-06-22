import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDataItem } from '../../../services/itemsService/IDataItem';

export interface IShowPageContextIssueState {
  items: IDataItem[];
  loading: boolean;
}
