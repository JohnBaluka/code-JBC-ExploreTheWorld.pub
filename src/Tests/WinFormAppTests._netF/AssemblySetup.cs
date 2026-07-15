// UI tests drive real windows in the foreground; running test classes in parallel
// makes them fight over focus and capture each other's screenshots.
[assembly: CollectionBehavior(DisableTestParallelization = true)]
