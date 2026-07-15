Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12960
    DatasheetFontHeight =11
    ItemSuffix =107
    Right =25290
    Bottom =17370
    TimerInterval =3000
    RecSrcDt = Begin
        0x3e84f21b898ae640
    End
    Caption ="Microsoft Excel - Watcher"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
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
            Height =540
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =11520
                    Top =120
                    Width =1320
                    Height =300
                    TabIndex =1
                    Name ="cmd_Connect"
                    Caption ="Connect"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =11520
                    LayoutCachedTop =120
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =420
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    Width =9960
                    Height =300
                    Name ="cmb_OpenFiles"
                    RowSourceType ="Value List"
                    OnChange ="[Event Procedure]"
                    HorizontalAnchor =2

                    LayoutCachedLeft =1500
                    LayoutCachedTop =120
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =120
                            Width =1320
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label108"
                            Caption ="File"
                            LayoutCachedLeft =60
                            LayoutCachedTop =120
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =420
                            ForeTint =100.0
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5520
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =420
                    Width =5640
                    Height =5040
                    Name ="txt_Log"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =5460
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
                    Left =4380
                    Top =60
                    Width =1320
                    Height =300
                    TabIndex =1
                    Name ="cmd_ClearLog"
                    Caption ="Clear Log"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =4380
                    LayoutCachedTop =60
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =360
                End
                Begin Subform
                    OverlapFlags =85
                    Left =5820
                    Top =420
                    Width =7080
                    Height =5040
                    TabIndex =2
                    Name ="Child105"
                    SourceObject ="Form.MsExcel_Events"
                    HorizontalAnchor =1
                    VerticalAnchor =2

                    LayoutCachedLeft =5820
                    LayoutCachedTop =420
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =5460
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5820
                            Top =60
                            Width =1320
                            Height =315
                            FontWeight =700
                            ForeColor =0
                            Name ="Label106"
                            Caption ="Log Events..."
                            HorizontalAnchor =1
                            LayoutCachedLeft =5820
                            LayoutCachedTop =60
                            LayoutCachedWidth =7140
                            LayoutCachedHeight =375
                            ForeTint =100.0
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =540
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
                    Left =9720
                    Top =120
                    Width =1500
                    Height =300
                    TabIndex =3
                    Name ="cmb_ExportMethod"
                    RowSourceType ="Value List"
                    RowSource ="COM;Direct"
                    DefaultValue ="COM"
                    HorizontalAnchor =1

                    LayoutCachedLeft =9720
                    LayoutCachedTop =120
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8160
                            Top =120
                            Width =1440
                            Height =300
                            ForeColor =0
                            Name ="Label1081"
                            Caption ="Export Method"
                            HorizontalAnchor =1
                            LayoutCachedLeft =8160
                            LayoutCachedTop =120
                            LayoutCachedWidth =9600
                            LayoutCachedHeight =420
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11280
                    Top =60
                    Width =1560
                    Height =420
                    Name ="cmd_SaveAsJson"
                    Caption ="Save As JSON"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =11280
                    LayoutCachedTop =60
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    Width =6600
                    Height =300
                    TabIndex =1
                    Name ="txt_OutputFilePath"
                    HorizontalAnchor =2

                    LayoutCachedLeft =1500
                    LayoutCachedTop =120
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =60
                            Top =120
                            Width =1320
                            Height =300
                            FontWeight =700
                            ForeColor =0
                            Name ="Label100"
                            Caption ="File"
                            LayoutCachedLeft =60
                            LayoutCachedTop =120
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =420
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =60
                    Top =120
                    Width =1320
                    Height =300
                    TabIndex =2
                    Name ="cmd_SelectFileOutput"
                    Caption ="Select..."
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =420
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "MsExcel.cls"
