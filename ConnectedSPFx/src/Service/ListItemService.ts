import { IListItem } from '../Contract/IListItems';
//import { sp, import { spODataEntityArray } from '@pnp/sp/odata';, Item } from "@pnp/sp"; 
import { spODataEntityArray } from '@pnp/sp/odata'
import { Item } from '@pnp/sp/items';
import { PnPClientStorage } from "@pnp/common";


export class ListitemService {
    private static storageKey: string = 'listItemsKey';
    private static listItems: IListItem[] = [];
    private static storage = new PnPClientStorage();


    public static getListItems(): Promise<IListItem[]> {
        return new Promise<IListItem[]>(async (resolve: (newListItems: IListItem[]) => void, reject: (error: any) => void) => {
            const loadedData = this.storage.session.get(this.storageKey);
            let items: IListItem[] ;//= null;
            if ((window as any).loadingData) {
                window.setTimeout((): void => {
                    this.getListItems().then((newListItems: IListItem[]): void => {
                        resolve(newListItems);
                    });
                }, 50);
            }
            else {
                if (!loadedData) {
                    (window as any).loadingData = true;
                    items = await sp.web.lists.getByTitle("Testlist").items.expand('Author').select("Id", "Title", "Author/Title").get(spODataEntityArray<Item, IListItem>(Item));
                    this.storage.session.put(this.storageKey, items);
                    console.info("from SharePoint list");
                    (window as any).loadingData = false;
                }
                else {
                    console.info("from Session storage");
                    items = loadedData;
                }
                resolve(items);
            }
        });
    }


}