Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17280
    DatasheetFontHeight =11
    ItemSuffix =97
    Right =17595
    Bottom =11790
    TimerInterval =3000
    RecSrcDt = Begin
        0xb6513b272f81e640
    End
    RecordSource ="cns_Country__FormQuery"
    Caption ="CountriesNow Data"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    SplitFormSize =10710
    SplitFormSize =10710
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri Light"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =780
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =15900
                    Top =60
                    Width =1320
                    Height =300
                    Name ="cmd_Load"
                    Caption ="Load"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =15900
                    LayoutCachedTop =60
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15900
                    Top =420
                    Width =1320
                    Height =300
                    TabIndex =1
                    Name ="cmd_Clear"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =15900
                    LayoutCachedTop =420
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =720
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =9060
                    Height =660
                    FontSize =24
                    FontWeight =700
                    Name ="Label96"
                    Caption ="Explore The World - CountriesNow.space API"
                    HyperlinkAddress ="https://countriesnow.space/"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =720
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =780
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =60
                    Width =1320
                    Height =300
                    Name ="Iso2"
                    ControlSource ="Iso2"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =60
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =1320
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label4"
                            Caption ="Iso2"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =360
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Top =60
                    Width =5640
                    Height =300
                    TabIndex =1
                    Name ="Country"
                    ControlSource ="Country"
                    HorizontalAnchor =2

                    LayoutCachedLeft =4380
                    LayoutCachedTop =60
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2940
                            Top =60
                            Width =1320
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label5"
                            Caption ="Country"
                            LayoutCachedLeft =2940
                            LayoutCachedTop =60
                            LayoutCachedWidth =4260
                            LayoutCachedHeight =360
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =420
                    Width =1320
                    Height =300
                    TabIndex =2
                    Name ="Iso3"
                    ControlSource ="Iso3"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =420
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =420
                            Width =1320
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label6"
                            Caption ="Iso3"
                            LayoutCachedLeft =60
                            LayoutCachedTop =420
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =720
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13020
                    Top =60
                    Width =4200
                    Height =300
                    TabIndex =3
                    Name ="GUID"
                    ControlSource ="GUID"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13020
                    LayoutCachedTop =60
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =12180
                            Top =60
                            Width =720
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label7"
                            Caption ="GUID"
                            HorizontalAnchor =1
                            LayoutCachedLeft =12180
                            LayoutCachedTop =60
                            LayoutCachedWidth =12900
                            LayoutCachedHeight =360
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13020
                    Top =420
                    Height =315
                    TabIndex =4
                    Name ="Row_ID"
                    ControlSource ="Row_ID"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13020
                    LayoutCachedTop =420
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =12120
                            Top =420
                            Width =780
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label95"
                            Caption ="Row ID"
                            HorizontalAnchor =1
                            LayoutCachedLeft =12120
                            LayoutCachedTop =420
                            LayoutCachedWidth =12900
                            LayoutCachedHeight =720
                            ForeTint =100.0
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =2880
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2460
                    Width =2400
                    Height =315
                    Name ="cmb_ExportType"
                    RowSourceType ="Value List"
                    RowSource ="Word;Excel;PowerPoint"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2460
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =2775
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2460
                            Width =1320
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label90"
                            Caption ="Export Type"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2460
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =2760
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5400
                    Top =2460
                    Width =8640
                    Height =315
                    TabIndex =1
                    Name ="txt_ExportFilePath"
                    HorizontalAnchor =2

                    LayoutCachedLeft =5400
                    LayoutCachedTop =2460
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =2775
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3960
                            Top =2460
                            Width =1320
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label91"
                            Caption ="Export File"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =2460
                            LayoutCachedWidth =5280
                            LayoutCachedHeight =2760
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14160
                    Top =2460
                    Height =315
                    TabIndex =2
                    Name ="cmd_SelectExportFile"
                    Caption ="Browse..."
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =14160
                    LayoutCachedTop =2460
                    LayoutCachedWidth =15600
                    LayoutCachedHeight =2775
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15660
                    Top =2460
                    Width =1560
                    Height =315
                    TabIndex =3
                    Name ="cmd_Export"
                    Caption ="Export"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =15660
                    LayoutCachedTop =2460
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =2775
                End
                Begin TextBox
                    TabStop = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =420
                    Width =17160
                    Height =1980
                    TabIndex =4
                    Name ="txt_Log"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =2400
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =60
                            Width =1320
                            Height =315
                            FontWeight =700
                            ForeColor =0
                            Name ="Label88"
                            Caption ="Log..."
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =375
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16080
                    Top =60
                    Width =1140
                    Height =315
                    TabIndex =5
                    Name ="cmd_ClearLog"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =16080
                    LayoutCachedTop =60
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =375
                End
            End
        End
    End
End
CodeBehindForm
' See "cns_Country.cls"
