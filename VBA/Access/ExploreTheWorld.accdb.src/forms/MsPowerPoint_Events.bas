Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6060
    DatasheetFontHeight =11
    ItemSuffix =102
    Left =10485
    Top =1725
    Right =17295
    Bottom =6495
    TimerInterval =3000
    RecSrcDt = Begin
        0xce3147a18d8ae640
    End
    RecordSource ="MsPowerPoint_Events_FormQuery"
    Caption ="MS PowerPoint Events"
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
        Begin Section
            CanGrow = NotDefault
            Height =1680
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CheckBox
                    OverlapFlags =85
                    Left =2040
                    Top =330
                    ColumnWidth =675
                    Name ="Log"
                    ControlSource ="Log"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =330
                    LayoutCachedWidth =2300
                    LayoutCachedHeight =570
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2270
                            Top =300
                            Width =390
                            Height =315
                            Name ="Label99"
                            Caption ="Log"
                            LayoutCachedLeft =2270
                            LayoutCachedTop =300
                            LayoutCachedWidth =2660
                            LayoutCachedHeight =615
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =720
                    Height =315
                    ColumnWidth =1770
                    TabIndex =1
                    Name ="Category"
                    ControlSource ="Category"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =720
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1035
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =720
                            Width =900
                            Height =315
                            Name ="Label100"
                            Caption ="Category"
                            LayoutCachedLeft =420
                            LayoutCachedTop =720
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1140
                    Height =315
                    ColumnWidth =4200
                    TabIndex =2
                    Name ="Name"
                    ControlSource ="Name"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1455
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =1140
                            Width =630
                            Height =315
                            Name ="Label101"
                            Caption ="Name"
                            LayoutCachedLeft =420
                            LayoutCachedTop =1140
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =1455
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MsPowerPoint_Events.cls"
