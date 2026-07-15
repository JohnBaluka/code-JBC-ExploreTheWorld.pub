-- ============================================================================
-- ExploreTheWorld Database Schema
-- Target: SQL Server (Database: ExploreTheWorld)
-- ============================================================================

USE [ExploreTheWorld];
GO

-- ============================================================================
-- RCC Tables (RestCountriesComApi) — parent tables first
-- ============================================================================

CREATE TABLE [dbo].[rcc_Country] (
    [Cca2]          NVARCHAR(2)     NOT NULL,
    [Cca3]          NVARCHAR(3)     NULL,
    [Region]        NVARCHAR(255)   NULL,
    [Subregion]     NVARCHAR(255)   NULL,
    [Population]    BIGINT          NOT NULL,
    CONSTRAINT [PK_rcc_Country] PRIMARY KEY ([Cca2])
);
GO

CREATE TABLE [dbo].[rcc_CountryName] (
    [Cca2]      NVARCHAR(2)     NOT NULL,
    [Common]    NVARCHAR(255)   NULL,
    [Official]  NVARCHAR(255)   NULL,
    CONSTRAINT [PK_rcc_CountryName] PRIMARY KEY ([Cca2]),
    CONSTRAINT [FK_rcc_CountryName_Cca2] FOREIGN KEY ([Cca2])
        REFERENCES [dbo].[rcc_Country] ([Cca2]) ON DELETE CASCADE
);
GO

CREATE TABLE [dbo].[rcc_CountryFlag] (
    [Cca2]  NVARCHAR(2)     NOT NULL,
    [Png]   NVARCHAR(MAX)   NULL,
    [Svg]   NVARCHAR(MAX)   NULL,
    [Alt]   NVARCHAR(MAX)   NULL,
    CONSTRAINT [PK_rcc_CountryFlag] PRIMARY KEY ([Cca2]),
    CONSTRAINT [FK_rcc_CountryFlag_Cca2] FOREIGN KEY ([Cca2])
        REFERENCES [dbo].[rcc_Country] ([Cca2]) ON DELETE CASCADE
);
GO

CREATE TABLE [dbo].[rcc_CountryCapital] (
    [ID]        INT             NOT NULL IDENTITY(1,1),
    [Cca2]      NVARCHAR(2)     NOT NULL,
    [Capital]   NVARCHAR(255)   NOT NULL,
    CONSTRAINT [PK_rcc_CountryCapital] PRIMARY KEY ([ID]),
    CONSTRAINT [FK_rcc_CountryCapital_Cca2] FOREIGN KEY ([Cca2])
        REFERENCES [dbo].[rcc_Country] ([Cca2]) ON DELETE CASCADE
);
GO

CREATE INDEX [IX_rcc_CountryCapital_Cca2] ON [dbo].[rcc_CountryCapital] ([Cca2]);
GO

CREATE TABLE [dbo].[rcc_CountryLanguage] (
    [ID]    INT             NOT NULL IDENTITY(1,1),
    [Cca2]  NVARCHAR(2)     NOT NULL,
    [Code]  NVARCHAR(10)    NOT NULL,
    [Name]  NVARCHAR(255)   NULL,
    CONSTRAINT [PK_rcc_CountryLanguage] PRIMARY KEY ([ID]),
    CONSTRAINT [FK_rcc_CountryLanguage_Cca2] FOREIGN KEY ([Cca2])
        REFERENCES [dbo].[rcc_Country] ([Cca2]) ON DELETE CASCADE,
    CONSTRAINT [UQ_rcc_CountryLanguage_Cca2_Code] UNIQUE ([Cca2], [Code])
);
GO

CREATE INDEX [IX_rcc_CountryLanguage_Cca2] ON [dbo].[rcc_CountryLanguage] ([Cca2]);
GO

CREATE TABLE [dbo].[rcc_CountryCurrency] (
    [ID]        INT             NOT NULL IDENTITY(1,1),
    [Cca2]      NVARCHAR(2)     NOT NULL,
    [Code]      NVARCHAR(10)    NOT NULL,
    [Name]      NVARCHAR(255)   NULL,
    [Symbol]    NVARCHAR(10)    NULL,
    CONSTRAINT [PK_rcc_CountryCurrency] PRIMARY KEY ([ID]),
    CONSTRAINT [FK_rcc_CountryCurrency_Cca2] FOREIGN KEY ([Cca2])
        REFERENCES [dbo].[rcc_Country] ([Cca2]) ON DELETE CASCADE,
    CONSTRAINT [UQ_rcc_CountryCurrency_Cca2_Code] UNIQUE ([Cca2], [Code])
);
GO

CREATE INDEX [IX_rcc_CountryCurrency_Cca2] ON [dbo].[rcc_CountryCurrency] ([Cca2]);
GO

-- ============================================================================
-- CNS Tables (CountriesNowSpaceApi) — parent tables first
-- ============================================================================

CREATE TABLE [dbo].[cns_Country] (
    [Iso2]      NVARCHAR(2)     NOT NULL,
    [Country]   NVARCHAR(255)   NULL,
    [Iso3]      NVARCHAR(3)     NULL,
    CONSTRAINT [PK_cns_Country] PRIMARY KEY ([Iso2])
);
GO

CREATE TABLE [dbo].[cns_City] (
    [ID]    INT             NOT NULL IDENTITY(1,1),
    [Iso2]  NVARCHAR(2)     NOT NULL,
    [City]  NVARCHAR(255)   NOT NULL,
    CONSTRAINT [PK_cns_City] PRIMARY KEY ([ID]),
    CONSTRAINT [FK_cns_City_Iso2] FOREIGN KEY ([Iso2])
        REFERENCES [dbo].[cns_Country] ([Iso2]) ON DELETE CASCADE
);
GO

CREATE INDEX [IX_cns_City_Iso2] ON [dbo].[cns_City] ([Iso2]);
GO

CREATE TABLE [dbo].[cns_CountryCapital] (
    [Iso2]      NVARCHAR(2)     NOT NULL,
    [Name]      NVARCHAR(255)   NULL,
    [Capital]   NVARCHAR(255)   NULL,
    CONSTRAINT [PK_cns_CountryCapital] PRIMARY KEY ([Iso2]),
    CONSTRAINT [FK_cns_CountryCapital_Iso2] FOREIGN KEY ([Iso2])
        REFERENCES [dbo].[cns_Country] ([Iso2]) ON DELETE CASCADE
);
GO

CREATE TABLE [dbo].[cns_CountryFlag] (
    [Iso2]      NVARCHAR(2)     NOT NULL,
    [Name]      NVARCHAR(255)   NULL,
    [Flag]      NVARCHAR(MAX)   NULL,
    [DialCode]  NVARCHAR(20)    NULL,
    CONSTRAINT [PK_cns_CountryFlag] PRIMARY KEY ([Iso2]),
    CONSTRAINT [FK_cns_CountryFlag_Iso2] FOREIGN KEY ([Iso2])
        REFERENCES [dbo].[cns_Country] ([Iso2]) ON DELETE CASCADE
);
GO

CREATE TABLE [dbo].[cns_CountryPopulation] (
    [ID]        INT             NOT NULL IDENTITY(1,1),
    [Country]   NVARCHAR(255)   NULL,
    [Code]      NVARCHAR(10)    NULL,
    [Iso3]      NVARCHAR(3)     NULL,
    CONSTRAINT [PK_cns_CountryPopulation] PRIMARY KEY ([ID])
);
GO

CREATE TABLE [dbo].[cns_PopulationCount] (
    [ID]                    INT     NOT NULL IDENTITY(1,1),
    [CountryPopulation_ID]  INT     NOT NULL,
    [Year]                  NVARCHAR(4)     NULL,
    [Value]                 BIGINT  NOT NULL,
    CONSTRAINT [PK_cns_PopulationCount] PRIMARY KEY ([ID]),
    CONSTRAINT [FK_cns_PopulationCount_CountryPopulation_ID] FOREIGN KEY ([CountryPopulation_ID])
        REFERENCES [dbo].[cns_CountryPopulation] ([ID]) ON DELETE CASCADE
);
GO

CREATE INDEX [IX_cns_PopulationCount_CountryPopulation_ID] ON [dbo].[cns_PopulationCount] ([CountryPopulation_ID]);
GO

CREATE TABLE [dbo].[cns_CountryStates] (
    [ID]    INT             NOT NULL IDENTITY(1,1),
    [Name]  NVARCHAR(255)   NULL,
    [Iso3]  NVARCHAR(3)     NULL,
    CONSTRAINT [PK_cns_CountryStates] PRIMARY KEY ([ID])
);
GO

CREATE TABLE [dbo].[cns_CountryState] (
    [ID]                INT             NOT NULL IDENTITY(1,1),
    [CountryStates_ID]  INT             NOT NULL,
    [Name]              NVARCHAR(255)   NULL,
    [StateCode]         NVARCHAR(10)    NULL,
    CONSTRAINT [PK_cns_CountryState] PRIMARY KEY ([ID]),
    CONSTRAINT [FK_cns_CountryState_CountryStates_ID] FOREIGN KEY ([CountryStates_ID])
        REFERENCES [dbo].[cns_CountryStates] ([ID]) ON DELETE CASCADE
);
GO

CREATE INDEX [IX_cns_CountryState_CountryStates_ID] ON [dbo].[cns_CountryState] ([CountryStates_ID]);
GO
