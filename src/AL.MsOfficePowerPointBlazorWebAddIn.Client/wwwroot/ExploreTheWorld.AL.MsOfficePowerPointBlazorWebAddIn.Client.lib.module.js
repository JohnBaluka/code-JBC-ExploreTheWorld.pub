console.log("Loading ExploreTheWorld.AL.MsOfficePowerPointBlazorWebAddIn.Client.lib.module.js");

export async function beforeWebAssemblyStart(options, extensions) {
    console.log("beforeWebAssemblyStart: hooking Office.onReady");
    Office.onReady((info) => {
        if (info.host === Office.HostType.PowerPoint) {
            console.log("Hosting in PowerPoint.");
            Office.addin.setStartupBehavior(Office.StartupBehavior.load);
        } else {
            console.log("Hosting in browser (not PowerPoint).");
        }
        console.log("Office onReady.");
    });
}

export async function beforeWebStart(options) { console.log("beforeWebStart"); }
export async function beforeServerStart(options, extensions) { console.log("beforeServerStart"); }
export async function afterWebStarted(blazor) { console.log("afterWebStarted"); }
export async function afterServerStarted(blazor) { console.log("afterServerStarted"); }
export async function afterWebAssemblyStarted(blazor) { console.log("afterWebAssemblyStarted"); }
