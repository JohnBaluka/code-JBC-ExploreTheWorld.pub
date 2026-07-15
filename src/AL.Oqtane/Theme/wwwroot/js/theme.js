// ExploreTheWorld Oqtane theme behaviors (plain JS, no Blazor interop).
// The theme renders statically in Oqtane's Static render mode, so Blazor @onclick
// handlers in the layout never fire. This delegated listener toggles Radzen's own
// sidebar classes instead and survives enhanced navigations.
(function () {
    if (window.etwThemeInitialized) return;
    window.etwThemeInitialized = true;

    document.addEventListener('click', function (e) {
        var toggle = e.target.closest('.rz-sidebar-toggle');
        if (!toggle) return;

        var sidebar = document.querySelector('.rz-layout > .rz-sidebar');
        if (!sidebar) return;

        sidebar.classList.toggle('rz-sidebar-expanded');
        sidebar.classList.toggle('rz-sidebar-collapsed');
    });
})();
