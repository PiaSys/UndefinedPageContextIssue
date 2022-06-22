import { IDataItem } from './IDataItem';

/**
 * Defines the abstract interface for the Items Service
 */
export interface IDataItemsService {

    /**
     * Returns the whole list of items
     * @returns The whole list of items
     */
    GetItems: () => Promise<IDataItem[]>;
}
