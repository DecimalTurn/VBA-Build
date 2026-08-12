Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6803
    DatasheetFontHeight =11
    ItemSuffix =3
    Right =24915
    Bottom =11730
    RecSrcDt = Begin
        0xfd5e93f4705de640
    End
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
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
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
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
        Begin Section
            Height =5669
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =566
                    Top =566
                    Width =2205
                    Height =675
                    FontSize =26
                    Name ="Label0"
                    Caption ="Test Form"
                    LayoutCachedLeft =566
                    LayoutCachedTop =566
                    LayoutCachedWidth =2771
                    LayoutCachedHeight =1241
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =1700
                    Width =2247
                    Height =561
                    Name ="cmdDevMode"
                    Caption ="Activate Dev Mode"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =566
                    LayoutCachedTop =1700
                    LayoutCachedWidth =2813
                    LayoutCachedHeight =2261
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2948
                    Top =1700
                    Width =3105
                    Height =525
                    FontSize =10
                    ForeColor =0
                    Name ="labDevMode"
                    Caption ="You must close and reopen \015\012the current database  to take effect"
                    LayoutCachedLeft =2948
                    LayoutCachedTop =1700
                    LayoutCachedWidth =6053
                    LayoutCachedHeight =2225
                    ForeTint =100.0
                End
            End
        End
    End
End
CodeBehindForm
' See "frmStart.cls"
