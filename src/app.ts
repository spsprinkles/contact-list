import { Dashboard, ItemForm } from "dattatable";
import { Components } from "gd-sprest-bs";
import { info } from "gd-sprest-bs/build/icons/svgs/info";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { personSquare } from "gd-sprest-bs/build/icons/svgs/personSquare";
import { plus } from "gd-sprest-bs/build/icons/svgs/plus";
import * as jQuery from "jquery";
import * as moment from "moment";
import { DataSource, IContactItem } from "./ds";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    // Constructor
    constructor(el: HTMLElement) {
        // Set the list name
        ItemForm.ListName = Strings.Lists.Contacts;

        // Render the dashboard
        this.render(el);
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        let dashboard = new Dashboard({
            el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [{
                    header: "By Status",
                    items: DataSource.StatusFilters,
                    onFilter: (value: string) => {
                        // Filter the table
                        dashboard.filter(4, value);
                    }
                }]
            },
            navigation: {
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "bg-sharepoint navbar-expand rounded-top";

                    // Set the brand
                    let brand = document.createElement("div");
                    brand.className = "d-flex";
                    brand.appendChild(personSquare());
                    brand.append(Strings.ProjectName);
                    brand.querySelector("svg").classList.add("me-75");
                    props.brand = brand;
                },
                // Adjust the brand alignment
                onRendered: (el) => {
                    el.querySelector("nav div.container-fluid").classList.add("ps-3");
                    el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("pe-none");
                },
                showFilter: false
            },
            subNavigation: {
                itemsEnd: [
                    {
                        text: "Add Contact",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: item.text,
                                btnProps: {
                                    // Render the icon button
                                    iconType: plus,
                                    iconSize: 28,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Create an item
                                        ItemForm.create({
                                            onUpdate: () => {
                                                // Load the data
                                                DataSource.load().then(items => {
                                                    // Refresh the table
                                                    dashboard.refresh(items);
                                                });
                                            }
                                        });
                                    }
                                },
                            });
                        }
                    }
                ],
                showFilter: true,
            },
            footer: {
                itemsEnd: [
                    {
                        className: "pe-none",
                        text: "v" + Strings.Version
                    }
                ]
            },
            table: {
                rows: DataSource.Items,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            "targets": 6,
                            "orderable": false,
                            "searchable": false
                        }
                    ],
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        let div = api.table().container() as HTMLDivElement;
                        let table = api.table().node() as HTMLTableElement;
                        div.querySelector(".dataTables_info").classList.add("text-center");
                        div.querySelector(".dataTables_length").classList.add("pt-2");
                        div.querySelector(".dataTables_paginate").classList.add("pt-03");
                        table.classList.remove("no-footer");
                        table.classList.add("tbl-footer");
                        table.classList.add("table-striped");
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Order by the 1st column by default; ascending
                    order: [[0, "asc"]]
                },
                columns: [
                    {
                        name: "Title",
                        title: "Employee Id"
                    },
                    {
                        name: "FullName",
                        title: "Name"
                    },
                    {
                        name: "PriPhone",
                        title: "Primary Phone"
                    },
                    {
                        name: "SecPhone",
                        title: "Secondary Phone"
                    },
                    {
                        name: "Status",
                        title: "Status"
                    },
                    {
                        name: "FinalEmploymentDate",
                        title: "Last Worked",
                        onRenderCell: (el, col, item: IContactItem) => {
                            if (item.FinalEmploymentDate) {
                                // Add the data-filter attribute for searching by date properly
                                el.setAttribute("data-filter", moment(item.FinalEmploymentDate).format("dddd MMMM DD YYYY"));
                                // Add the data-order attribute for sorting by date properly
                                el.setAttribute("data-order", item.FinalEmploymentDate);
                                el.innerText = moment(item.FinalEmploymentDate).format(Strings.DateFormat);
                            }
                        }
                    },
                    {
                        className: "text-end text-nowrap",
                        name: "Actions",
                        isHidden: true,
                        onRenderCell: (el, col, item: IContactItem) => {
                            // Create a span to wrap the icons in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            let spanEdit = span.cloneNode() as HTMLSpanElement;
                            spanEdit.classList.add("me-1");

                            // Add the view item icon
                            el.appendChild(span);

                            // Add the edit item icon
                            el.appendChild(spanEdit);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: "View Info",
                                btnProps: {
                                    // Render the icon button
                                    iconType: info,
                                    iconSize: 28,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the display form
                                        ItemForm.view({
                                            itemId: item.Id
                                        });
                                    }
                                },
                            });

                            // Render a tooltip
                            Components.Tooltip({
                                el: spanEdit,
                                content: "Edit Info",
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1",
                                    iconType: pencilSquare,
                                    iconSize: 24,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the edit form
                                        ItemForm.edit({
                                            itemId: item.Id,
                                            onUpdate: () => {
                                                // Refresh the data
                                                DataSource.load().then(items => {
                                                    // Update the data
                                                    dashboard.refresh(items);
                                                });
                                            }
                                        });
                                    }
                                },
                            });
                        }
                    }
                ]
            }
        });
    }
}