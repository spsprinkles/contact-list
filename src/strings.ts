import { ContextInfo, Helper } from "gd-sprest-bs";

// Global Path
let AssetsUrl: string = ContextInfo.webServerRelativeUrl + "/SiteAssets/";
let SPFxContext: { pageContext: any; sdks: { microsoftTeams: any } } = null;

// Sets the context information for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    SPFxContext = context;
    ContextInfo.setPageContext(SPFxContext.pageContext);

    // Load the default scripts
    Helper.loadSPCore();

    // Set the teams flag
    Strings.IsTeams = SPFxContext.sdks.microsoftTeams ? true : false;

    // Update the global path
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
    AssetsUrl = Strings.SourceUrl + "/SiteAssets/";

    // Update the solution url
    Strings.SolutionUrl = AssetsUrl + "index.html";
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "contact-list",
    DateFormat: "D-MMM-YYYY",
    GlobalVariable: "ContactList",
    IsTeams: false,
    Lists: {
        Contacts: "Contacts"
    },
    ProjectName: "Contact List",
    ProjectDescription: "Easily manage a list of contacts.",
    SolutionUrl: AssetsUrl + "index.html",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;