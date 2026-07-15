export function IsRunningInHost() {
    return typeof Office !== "undefined" &&
        typeof Office.context !== "undefined" &&
        Office.context !== null;
}
