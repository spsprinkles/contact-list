import { Components, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

// Item
export interface IContactItem extends Types.SP.ListItem {
    FirstName: string;
    LastName: string;
    FullName: string;
    Address: string;
    PriPhone: string;
    SecPhone: string;
    Title: string;
    Status: string;
    FinalEmploymentDate: string;
}

/**
 * Data Source
 */
export class DataSource {
    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters(): PromiseLike<Components.ICheckboxGroupItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the status field
            Web(Strings.SourceUrl).Lists(Strings.Lists.Contacts).Fields("Status").execute((fld: Types.SP.FieldChoice) => {
                let items: Components.ICheckboxGroupItem[] = [];

                // Parse the choices
                for (let i = 0; i < fld.Choices.results.length; i++) {
                    // Add an item
                    items.push({
                        label: fld.Choices.results[i],
                        type: Components.CheckboxGroupTypes.Switch
                    });
                }

                // Set the filters and resolve the promise
                this._statusFilters = items;
                resolve(items);
            }, reject);
        });
    }

    // Gets the item id from the query string
    static getItemIdFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0];
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "ID") {
                // Return the item
                return parseInt(value);
            }
        }
    }

    // Initializes the application
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            this.load().then(() => {
                // Load the status filters
                this.loadStatusFilters().then(() => {
                    // Resolve the request
                    resolve();
                }, reject);
            }, reject)
        });
    }

    // Loads the list data
    private static _items: IContactItem[] = null;
    static get Items(): IContactItem[] { return this._items; }
    static load(): PromiseLike<IContactItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            Web(Strings.SourceUrl).Lists(Strings.Lists.Contacts).Items().query({
                GetAllItems: true,
                OrderBy: ["Title"],
                Top: 5000
            }).execute(
                // Success
                items => {
                    // Set the items
                    this._items = items.results as any;

                    // Resolve the request
                    resolve(this._items);
                },
                // Error
                () => { reject(); }
            );
        });
    }
}