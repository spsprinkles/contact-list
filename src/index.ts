import { InstallationRequired } from "dattatable";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el, context?, sourceUrl?: string) => {
        // See if the SPFx page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);
        }
        
        // Hide the first column of the webpart zones
        let wpZone: HTMLElement = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td");
        wpZone ? wpZone.style.width = "0%" : null;

        // Make the second column of the webpart zones full width
        wpZone = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td:last-child");
        wpZone ? wpZone.style.width = "100%" : null;

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Create the application
                new App(el);
            },

            // Error
            () => {
                // See if an installation is required
                InstallationRequired.requiresInstall(Configuration).then(installFl => {
                    // See if an install is required
                    if (installFl) {
                        // Show the dialog
                        InstallationRequired.showDialog();
                    } else {
                        // Log
                        console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
                    }
                });
            }
        );
    }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;

    // Render the application
    GlobalVariable.render(elApp);
}