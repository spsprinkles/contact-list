import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                AllowContentTypes: true,
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                ContentTypesEnabled: true,
                Description: "A list of contacts with details",
                EnableAttachments: false,
                EnableFolderCreation: false,
                OnQuickLaunch: false,
                Title: Strings.Lists.Contacts
            },
            ContentTypes: [
                {
                    Name: "Contact",
                    Description: "Contact Information",
                    FieldRefs: [
                        "FirstName",
                        "LastName",
                        "FullName",
                        "Address",
                        "PriPhone",
                        "SecPhone",
                        "Title",
                        "Status",
                        "FinalEmploymentDate"
                    ]
                }
            ],
            TitleFieldDisplayName: "Employee Id",
            CustomFields: [
                {
                    name: "FirstName",
                    title: "First Name",
                    type: Helper.SPCfgFieldType.Text,
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "LastName",
                    title: "Last Name",
                    type: Helper.SPCfgFieldType.Text,
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "FullName",
                    title: "Name",
                    type: Helper.SPCfgFieldType.Calculated,
                    fieldRefs: ["FirstName","LastName"],
                    formula: "=CONCATENATE([First Name],\" \",[Last Name])",
                    resultType: SPTypes.FieldResultType.Text
                } as Helper.IFieldInfoCalculated,
                {
                    name: "Address",
                    title: "Address",
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 3,
                    type: Helper.SPCfgFieldType.Note
                } as Helper.IFieldInfoNote,
                {
                    name: "PriPhone",
                    title: "Primary Phone",
                    type: Helper.SPCfgFieldType.Text,
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "SecPhone",
                    title: "Secondary Phone",
                    type: Helper.SPCfgFieldType.Text
                } as Helper.IFieldInfoText,
                {
                    name: "Status",
                    title: "Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "Active",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    indexed: true,
                    required: true,
                    showInNewForm: false,
                    choices: [
                        "Active", "Leave of absence", "Retired", "Terminated"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "FinalEmploymentDate",
                    title: "Last Worked",
                    type: Helper.SPCfgFieldType.Date,
                    displayFormat: SPTypes.DateFormat.DateOnly,
                    showInNewForm: false,
                } as Helper.IFieldInfoDate
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "FirstName", "LastName", "FullName", "Address", "PriPhone", "SecPhone", "Status", "FinalEmploymentDate"
                    ]
                }
            ]
        }
    ]
});

// Adds the solution to a classic page
Configuration["addToPage"] = (pageUrl: string) => {
    // Add a content editor webpart to the page
    Helper.addContentEditorWebPart(pageUrl, {
        contentLink: Strings.SolutionUrl,
        description: Strings.ProjectDescription,
        frameType: "None",
        title: Strings.ProjectName
    }).then(
        // Success
        () => {
            // Load
            console.log("[" + Strings.ProjectName + "] Successfully added the solution to the page.", pageUrl);
        },

        // Error
        ex => {
            // Load
            console.log("[" + Strings.ProjectName + "] Error adding the solution to the page.", ex);
        }
    );
}