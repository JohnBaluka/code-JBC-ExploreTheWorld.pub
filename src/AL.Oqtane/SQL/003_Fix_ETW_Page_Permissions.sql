-- 003_Fix_ETW_Page_Permissions.sql
-- Fixes the orphaned permission rows inserted by 002_Create_ETW_Pages.sql.
-- Root cause: 002 used 'Everyone'/'Admin' but actual Oqtane RoleNames are
--             'All Users' / 'Administrators'.
-- This script deletes the bad NULL-RoleId permission rows and reinserts them
-- with the correct RoleIds, then does the same for the Module permissions.
-- Idempotent: safe to run more than once.
-- Run against the Oqtane TENANT database (DefaultConnection).

DECLARE @Now  DATETIME2     = GETDATE();
DECLARE @By   NVARCHAR(255) = 'system';

DECLARE @SiteId INT = (SELECT TOP 1 SiteId FROM [Site] ORDER BY SiteId);

-- Correct role IDs
DECLARE @RoleId_Everyone INT = (
    SELECT TOP 1 RoleId FROM [Role]
    WHERE [Name] = 'All Users' AND IsSystem = 1);

DECLARE @RoleId_Admin INT = (
    SELECT TOP 1 RoleId FROM [Role]
    WHERE [Name] = 'Administrators' AND SiteId = @SiteId);

PRINT 'SiteId=' + CAST(@SiteId AS NVARCHAR)
    + '  RoleId_Everyone=' + ISNULL(CAST(@RoleId_Everyone AS NVARCHAR), 'NULL')
    + '  RoleId_Admin='    + ISNULL(CAST(@RoleId_Admin    AS NVARCHAR), 'NULL');

-- Page IDs we need to fix
DECLARE @PageId_CN INT = (SELECT TOP 1 PageId FROM [Page]
    WHERE SiteId = @SiteId AND [Path] = 'countries-now' AND IsDeleted = 0);
DECLARE @PageId_CS INT = (SELECT TOP 1 PageId FROM [Page]
    WHERE SiteId = @SiteId AND [Path] = 'country-slides' AND IsDeleted = 0);

-- Module IDs (from PageModule, assuming one module per page in Default pane)
DECLARE @ModuleId_CN INT = (
    SELECT TOP 1 pm.ModuleId FROM [PageModule] pm
    WHERE pm.PageId = @PageId_CN AND pm.IsDeleted = 0);
DECLARE @ModuleId_CS INT = (
    SELECT TOP 1 pm.ModuleId FROM [PageModule] pm
    WHERE pm.PageId = @PageId_CS AND pm.IsDeleted = 0);

PRINT 'PageId_CN=' + ISNULL(CAST(@PageId_CN AS NVARCHAR), 'NULL')
    + '  ModuleId_CN=' + ISNULL(CAST(@ModuleId_CN AS NVARCHAR), 'NULL');
PRINT 'PageId_CS=' + ISNULL(CAST(@PageId_CS AS NVARCHAR), 'NULL')
    + '  ModuleId_CS=' + ISNULL(CAST(@ModuleId_CS AS NVARCHAR), 'NULL');

-- ─── Delete the orphaned NULL-RoleId permissions ───────────────────────────
-- (rows that have RoleId IS NULL AND UserId IS NULL are never used by Oqtane)
DELETE FROM [Permission]
WHERE SiteId = @SiteId
  AND EntityName IN ('Page', 'Module')
  AND EntityId IN (@PageId_CN, @PageId_CS, @ModuleId_CN, @ModuleId_CS)
  AND RoleId IS NULL
  AND UserId IS NULL;

PRINT 'Deleted orphaned permission rows: ' + CAST(@@ROWCOUNT AS NVARCHAR);

-- ─── Reinsert Page permissions ─────────────────────────────────────────────
IF @PageId_CN IS NOT NULL
BEGIN
    INSERT INTO [Permission]
        (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Page', @PageId_CN, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Page', @PageId_CN, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);
    PRINT 'Inserted page permissions for countries-now';
END

IF @PageId_CS IS NOT NULL
BEGIN
    INSERT INTO [Permission]
        (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Page', @PageId_CS, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Page', @PageId_CS, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);
    PRINT 'Inserted page permissions for country-slides';
END

-- ─── Reinsert Module permissions ───────────────────────────────────────────
IF @ModuleId_CN IS NOT NULL
BEGIN
    INSERT INTO [Permission]
        (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Module', @ModuleId_CN, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Module', @ModuleId_CN, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);
    PRINT 'Inserted module permissions for CountriesNow';
END

IF @ModuleId_CS IS NOT NULL
BEGIN
    INSERT INTO [Permission]
        (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Module', @ModuleId_CS, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Module', @ModuleId_CS, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);
    PRINT 'Inserted module permissions for CountrySlides';
END
