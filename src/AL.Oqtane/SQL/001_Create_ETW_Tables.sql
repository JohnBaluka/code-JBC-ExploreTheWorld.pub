-- ExploreTheWorld — CountriesNow.space API cache tables
-- Applies to: Oqtane-ETW.mdf (SQL Server LocalDB, DefaultConnection in appsettings.json)
-- These tables are auto-created by EF Core EnsureCreated() on first run.
-- This script is provided as a reference and for manual schema recreation if needed.
-- Run against the Oqtane-ETW database (not the Oqtane master/tenant DB).

-- ============================================================
-- cns_Country — country list cached from CountriesNow.space API
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_Country')
BEGIN
    CREATE TABLE [dbo].[cns_Country] (
        [Iso2]    NVARCHAR(10)  NOT NULL,
        [Country] NVARCHAR(200) NULL,
        [Iso3]    NVARCHAR(10)  NULL,
        CONSTRAINT [PK_cns_Country] PRIMARY KEY CLUSTERED ([Iso2])
    );
END;
GO

-- ============================================================
-- cns_City — cities per country
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_City')
BEGIN
    CREATE TABLE [dbo].[cns_City] (
        [ID]   INT           NOT NULL IDENTITY(1,1),
        [Iso2] NVARCHAR(10)  NOT NULL DEFAULT '',
        [City] NVARCHAR(200) NOT NULL DEFAULT '',
        CONSTRAINT [PK_cns_City] PRIMARY KEY CLUSTERED ([ID]),
        CONSTRAINT [FK_cns_City_cns_Country] FOREIGN KEY ([Iso2])
            REFERENCES [dbo].[cns_Country] ([Iso2])
    );
    CREATE NONCLUSTERED INDEX [IX_cns_City_Iso2] ON [dbo].[cns_City] ([Iso2]);
END;
GO

-- ============================================================
-- cns_CountryCapital — capital city per country (1:1 with cns_Country)
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_CountryCapital')
BEGIN
    CREATE TABLE [dbo].[cns_CountryCapital] (
        [Iso2]    NVARCHAR(10)  NOT NULL DEFAULT '',
        [Name]    NVARCHAR(200) NULL,
        [Capital] NVARCHAR(200) NULL,
        CONSTRAINT [PK_cns_CountryCapital] PRIMARY KEY CLUSTERED ([Iso2]),
        CONSTRAINT [FK_cns_CountryCapital_cns_Country] FOREIGN KEY ([Iso2])
            REFERENCES [dbo].[cns_Country] ([Iso2])
    );
END;
GO

-- ============================================================
-- cns_CountryFlag — flag URL and dial code per country (1:1 with cns_Country)
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_CountryFlag')
BEGIN
    CREATE TABLE [dbo].[cns_CountryFlag] (
        [Iso2]     NVARCHAR(10)   NOT NULL DEFAULT '',
        [Name]     NVARCHAR(200)  NULL,
        [Flag]     NVARCHAR(2000) NULL,
        [DialCode] NVARCHAR(20)   NULL,
        CONSTRAINT [PK_cns_CountryFlag] PRIMARY KEY CLUSTERED ([Iso2]),
        CONSTRAINT [FK_cns_CountryFlag_cns_Country] FOREIGN KEY ([Iso2])
            REFERENCES [dbo].[cns_Country] ([Iso2])
    );
END;
GO

-- ============================================================
-- cns_CountryPopulation — population records per country
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_CountryPopulation')
BEGIN
    CREATE TABLE [dbo].[cns_CountryPopulation] (
        [ID]      INT           NOT NULL IDENTITY(1,1),
        [Country] NVARCHAR(200) NULL,
        [Code]    NVARCHAR(20)  NULL,
        [Iso3]    NVARCHAR(10)  NULL,
        CONSTRAINT [PK_cns_CountryPopulation] PRIMARY KEY CLUSTERED ([ID])
    );
END;
GO

-- ============================================================
-- cns_PopulationCount — annual counts belonging to cns_CountryPopulation
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_PopulationCount')
BEGIN
    CREATE TABLE [dbo].[cns_PopulationCount] (
        [ID]                    INT           NOT NULL IDENTITY(1,1),
        [CountryPopulation_ID]  INT           NOT NULL,
        [Year]                  NVARCHAR(10)  NULL,
        [Value]                 BIGINT        NOT NULL DEFAULT 0,
        CONSTRAINT [PK_cns_PopulationCount] PRIMARY KEY CLUSTERED ([ID]),
        CONSTRAINT [FK_cns_PopulationCount_cns_CountryPopulation] FOREIGN KEY ([CountryPopulation_ID])
            REFERENCES [dbo].[cns_CountryPopulation] ([ID])
    );
    CREATE NONCLUSTERED INDEX [IX_cns_PopulationCount_CountryPopulation_ID]
        ON [dbo].[cns_PopulationCount] ([CountryPopulation_ID]);
END;
GO

-- ============================================================
-- cns_CountryStates — states/provinces grouping per country
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_CountryStates')
BEGIN
    CREATE TABLE [dbo].[cns_CountryStates] (
        [ID]   INT           NOT NULL IDENTITY(1,1),
        [Name] NVARCHAR(200) NULL,
        [Iso3] NVARCHAR(10)  NULL,
        CONSTRAINT [PK_cns_CountryStates] PRIMARY KEY CLUSTERED ([ID])
    );
END;
GO

-- ============================================================
-- cns_CountryState — individual state/province rows
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'cns_CountryState')
BEGIN
    CREATE TABLE [dbo].[cns_CountryState] (
        [ID]               INT           NOT NULL IDENTITY(1,1),
        [CountryStates_ID] INT           NOT NULL,
        [Name]             NVARCHAR(200) NULL,
        [StateCode]        NVARCHAR(20)  NULL,
        CONSTRAINT [PK_cns_CountryState] PRIMARY KEY CLUSTERED ([ID]),
        CONSTRAINT [FK_cns_CountryState_cns_CountryStates] FOREIGN KEY ([CountryStates_ID])
            REFERENCES [dbo].[cns_CountryStates] ([ID])
    );
    CREATE NONCLUSTERED INDEX [IX_cns_CountryState_CountryStates_ID]
        ON [dbo].[cns_CountryState] ([CountryStates_ID]);
END;
GO
