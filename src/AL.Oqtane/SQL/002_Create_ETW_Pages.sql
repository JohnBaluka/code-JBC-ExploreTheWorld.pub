-- 002_Create_ETW_Pages.sql
-- Inserts the Countries Now and Country Slides pages with module instances and permissions.
-- Idempotent: skips each page if its path already exists (non-deleted).
-- Run against the Oqtane TENANT database (DefaultConnection).

DECLARE @Now  DATETIME2     = GETDATE();
DECLARE @By   NVARCHAR(255) = 'system';

-- First site (single-tenant setup)
DECLARE @SiteId INT = (SELECT TOP 1 SiteId FROM [Site] ORDER BY SiteId);

-- Role IDs for permission grants
-- Note: Oqtane RoleNames.Everyone = 'All Users', RoleNames.Admin = 'Administrators'
DECLARE @RoleId_Everyone INT = (SELECT TOP 1 RoleId FROM [Role]
    WHERE [Name] = 'All Users' AND IsSystem = 1);
DECLARE @RoleId_Admin INT = (SELECT TOP 1 RoleId FROM [Role]
    WHERE [Name] = 'Administrators' AND SiteId = @SiteId);

-- Assembly-qualified module definition names (namespace, assemblyName)
DECLARE @MDN_CN NVARCHAR(500) =
    'JBC.ExploreTheWorld.AL.Oqtane.CountriesNow, JBC.ExploreTheWorld.AL.Oqtane.CountriesNow__Module.Client';
DECLARE @MDN_CS NVARCHAR(500) =
    'JBC.ExploreTheWorld.AL.Oqtane.CountrySlides, JBC.ExploreTheWorld.AL.Oqtane.CountrySlides__Module.Client';

-- ─────────────────────────────────────────────────────────────────
-- 1. Countries Now  (path: countries-now)
-- ─────────────────────────────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM [Page]
               WHERE SiteId = @SiteId AND [Path] = 'countries-now' AND IsDeleted = 0)
BEGIN
    INSERT INTO [Page]
        (SiteId, [Path], ParentId, [Name], Title, [Order],
         Url, ThemeType, DefaultContainerType, HeadContent, BodyContent,
         Icon, IsNavigation, IsClickable, IsPersonalizable, UserId,
         EffectiveDate, ExpiryDate, IsDeleted, DeletedBy, DeletedOn,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'countries-now', NULL, 'Countries Now', 'Countries Now', 11,
         '', '', '', '', '',
         '', 1, 1, 0, NULL,
         NULL, NULL, 0, NULL, NULL,
         @By, @Now, @By, @Now);

    DECLARE @PageId_CN INT = SCOPE_IDENTITY();

    -- Page permissions
    INSERT INTO [Permission] (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Page', @PageId_CN, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Page', @PageId_CN, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);

    -- Module instance (Module table has no IsDeleted/DeletedBy/DeletedOn — they are [NotMapped])
    INSERT INTO [Module] (SiteId, ModuleDefinitionName, AllPages, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES (@SiteId, @MDN_CN, 0, @By, @Now, @By, @Now);

    DECLARE @ModuleId_CN INT = SCOPE_IDENTITY();

    -- Module permissions
    INSERT INTO [Permission] (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Module', @ModuleId_CN, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Module', @ModuleId_CN, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);

    -- Place module on page in the Default pane
    INSERT INTO [PageModule]
        (PageId, ModuleId, Title, Pane, [Order], ContainerType,
         EffectiveDate, ExpiryDate, Header, Footer,
         IsDeleted, DeletedBy, DeletedOn,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@PageId_CN, @ModuleId_CN, 'Countries Now', 'Default', 1, '',
         NULL, NULL, '', '',
         0, NULL, NULL,
         @By, @Now, @By, @Now);

    PRINT 'Created page: countries-now (PageId=' + CAST(@PageId_CN AS NVARCHAR) + ', ModuleId=' + CAST(@ModuleId_CN AS NVARCHAR) + ')';
END
ELSE
    PRINT 'Skipped: countries-now already exists';

-- ─────────────────────────────────────────────────────────────────
-- 2. Country Slides  (path: country-slides)
-- ─────────────────────────────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM [Page]
               WHERE SiteId = @SiteId AND [Path] = 'country-slides' AND IsDeleted = 0)
BEGIN
    INSERT INTO [Page]
        (SiteId, [Path], ParentId, [Name], Title, [Order],
         Url, ThemeType, DefaultContainerType, HeadContent, BodyContent,
         Icon, IsNavigation, IsClickable, IsPersonalizable, UserId,
         EffectiveDate, ExpiryDate, IsDeleted, DeletedBy, DeletedOn,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'country-slides', NULL, 'Country Slides', 'Country Slides', 12,
         '', '', '', '', '',
         '', 1, 1, 0, NULL,
         NULL, NULL, 0, NULL, NULL,
         @By, @Now, @By, @Now);

    DECLARE @PageId_CS INT = SCOPE_IDENTITY();

    -- Page permissions
    INSERT INTO [Permission] (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Page', @PageId_CS, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Page', @PageId_CS, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);

    -- Module instance
    INSERT INTO [Module] (SiteId, ModuleDefinitionName, AllPages, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES (@SiteId, @MDN_CS, 0, @By, @Now, @By, @Now);

    DECLARE @ModuleId_CS INT = SCOPE_IDENTITY();

    -- Module permissions
    INSERT INTO [Permission] (SiteId, EntityName, EntityId, PermissionName, RoleId, UserId, IsAuthorized, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@SiteId, 'Module', @ModuleId_CS, 'View', @RoleId_Everyone, NULL, 1, @By, @Now, @By, @Now),
        (@SiteId, 'Module', @ModuleId_CS, 'Edit', @RoleId_Admin,    NULL, 1, @By, @Now, @By, @Now);

    -- Place module on page
    INSERT INTO [PageModule]
        (PageId, ModuleId, Title, Pane, [Order], ContainerType,
         EffectiveDate, ExpiryDate, Header, Footer,
         IsDeleted, DeletedBy, DeletedOn,
         CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
    VALUES
        (@PageId_CS, @ModuleId_CS, 'Country Slides', 'Default', 1, '',
         NULL, NULL, '', '',
         0, NULL, NULL,
         @By, @Now, @By, @Now);

    PRINT 'Created page: country-slides (PageId=' + CAST(@PageId_CS AS NVARCHAR) + ', ModuleId=' + CAST(@ModuleId_CS AS NVARCHAR) + ')';
END
ELSE
    PRINT 'Skipped: country-slides already exists';
