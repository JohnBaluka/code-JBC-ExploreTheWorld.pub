-- 004_Set_ETW_Theme.sql
-- Makes the ExploreTheWorld Radzen theme the site default (layout + container).
-- The theme assembly registers itself in the [Theme] table on startup; this script
-- only switches the site defaults to it. Page rows keep ThemeType = '' so they
-- inherit the site default.
-- Idempotent: safe to run more than once.
-- Run against the Oqtane TENANT database (DefaultConnection).

SET QUOTED_IDENTIFIER ON; -- required: [Setting] has a filtered index; sqlcmd defaults to OFF

DECLARE @Now DATETIME2     = GETDATE();
DECLARE @By  NVARCHAR(255) = 'system';

-- First site (single-tenant setup)
DECLARE @SiteId INT = (SELECT TOP 1 SiteId FROM [Site] ORDER BY SiteId);

DECLARE @ThemeType NVARCHAR(200) =
    'JBC.ExploreTheWorld.AL.Oqtane.Theme.Default, JBC.ExploreTheWorld.AL.Oqtane.Theme';
DECLARE @ContainerType NVARCHAR(200) =
    'JBC.ExploreTheWorld.AL.Oqtane.Theme.Container, JBC.ExploreTheWorld.AL.Oqtane.Theme';

UPDATE [Site]
SET DefaultThemeType     = @ThemeType,
    DefaultContainerType = @ContainerType,
    ModifiedBy           = @By,
    ModifiedOn           = @Now
WHERE SiteId = @SiteId
  AND (DefaultThemeType <> @ThemeType OR DefaultContainerType <> @ContainerType);

PRINT 'Site ' + CAST(@SiteId AS NVARCHAR) + ' default theme set to ExploreTheWorld Radzen Theme ('
    + CAST(@@ROWCOUNT AS NVARCHAR) + ' row(s) updated).';

-- Normalize legacy "-base.css" RadzenCss settings saved before the switch to the
-- complete Radzen theme files (default.css, dark.css, ...).
UPDATE [Setting]
SET SettingValue = REPLACE(SettingValue, '-base.css', '.css'),
    ModifiedBy   = @By,
    ModifiedOn   = @Now
WHERE EntityName   = 'Site'
  AND SettingName  = 'JBC.ExploreTheWorld.AL.Oqtane.Theme:RadzenCss'
  AND SettingValue LIKE '%-base.css';

PRINT 'RadzenCss setting normalized (' + CAST(@@ROWCOUNT AS NVARCHAR) + ' row(s) updated).';
