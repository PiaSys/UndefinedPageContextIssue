import { IDataItem } from './IDataItem';
import { IDataItemsService } from './IDataItemsService';

// Import PnPjs types
import { SPFI } from '@pnp/sp';
import { Caching } from "@pnp/queryable";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { ISiteUser } from '@pnp/sp/site-users/types';
import { IList, ICamlQuery } from "@pnp/sp/lists";

import { SPHttpClient, SPHttpClientResponse, MSGraphClient, AadHttpClient, HttpClientResponse, AadHttpClientConfiguration } from '@microsoft/sp-http';

/**
 * Defines the concrete implementation of the interface for the Assets Service
 */
export class DataItemsService implements IDataItemsService {

    private _sp: SPFI = null;
    private _aadClient: AadHttpClient = null;
    private _sourceList: string = null;

    constructor(sp: SPFI, aadHttpClient: AadHttpClient, sourceList: string) {
        this._sp = sp;
        this._aadClient = aadHttpClient;
        this._sourceList = sourceList;
    }

    /**
     * Returns the whole list of assets
     * @returns The whole list of assets
     */
    public async GetItems(): Promise<IDataItem[]> {

        const list: IList = this._sp.web.lists.getById(this._sourceList);

        // Build the CAML query
        const caml: ICamlQuery = {
            ViewXml: `<View><ViewFields><FieldRef Name='ID' /><FieldRef Name='Title' /></ViewFields><RowLimit>10</RowLimit></View>`,
        };

        // Get list items
        const queryResult = await list.getItemsByCAMLQuery(caml);

        const result: IDataItem[] = [];

        // Map the query results to the actual data items
        for (const i of queryResult) {

            const editor: ISiteUser = await this._sp.web.getUserById(i["EditorId"])
                .select('Id', 'Email', 'LoginName', 'Title')();

            const editorUPN: string = editor['LoginName'].substring(editor['LoginName'].lastIndexOf('|') + 1);
            const editorAvatar: string = await this.getUserPicture(editorUPN);

            const newItem: IDataItem = {
                id: i['ID'],
                title: i['Title'],
                name: editor['Title'],
                profileImageSrc: editorAvatar,
                location: "SharePoint",
                activity: i['Modified']
            };

            result.push(newItem);
        }

        return result;
    }

    private async getUserPicture(upn: string): Promise<string> {

        const photoResponse: HttpClientResponse = await this._aadClient.get(
            `https://graph.microsoft.com/v1.0/users/${upn}/photo/$value`,
            AadHttpClient.configurations.v1
        );

        const photo: string = await this.blobToBase64(await photoResponse.blob());

        return photo;
    }

    private async blobToBase64(blob: Blob): Promise<string> {
        return new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onerror = reject;
          reader.onload = _ => {
            resolve(reader.result as string);
          };
          reader.readAsDataURL(blob);
        });
    }
}