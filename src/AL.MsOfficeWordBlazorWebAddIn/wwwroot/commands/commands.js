// commands.js — Shared Runtime add-in command entry point.
// Implements toggleTaskpane to show/hide the task pane.

let taskpaneVisible = false;

function toggleTaskpane(event) {
    if (taskpaneVisible) {
        Office.addin.hide();
        taskpaneVisible = false;
    } else {
        Office.addin.showAsTaskpane();
        taskpaneVisible = true;
    }
    event.completed();
}

Office.actions.associate("toggleTaskpane", toggleTaskpane);
