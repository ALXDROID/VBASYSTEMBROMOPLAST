Version =21
VersionRequired =20
PublishOption =1
Checksum =696390575
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10512
    DatasheetFontHeight =11
    ItemSuffix =3
    Left =6192
    Top =1536
    Right =16908
    Bottom =11784
    DatasheetGridlinesColor =15461355
    RecSrcDt = Begin
        0xa42bd7ce1a49e640
    End
    GUID = Begin
        0xac63a6b97009d54ca3300f44d87edf74
    End
    NameMap = Begin
        0x0acc0e5500000000000000000000000000000000000000000c00000005000000 ,
        0x0000000000000000000000000000
    End
    DatasheetFontName ="Trebuchet MS"
    OnMouseDown ="[Event Procedure]"
    AllowDatasheetView =0
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
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            BorderColor =8355711
            ForeColor =6710886
            FontName ="Trebuchet MS"
            GridlineColor =10921638
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =10921638
            GridlineColor =10921638
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            BorderColor =10921638
            GridlineColor =10921638
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            BorderColor =10921638
            ForeColor =2235926
            GridlineColor =10921638
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7272
            Name ="Detail"
            GUID = Begin
                0xe88379da1506674db626c846ff5ed5e8
            End
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Image
                    BackStyle =1
                    BorderWidth =3
                    SizeMode =1
                    PictureType =2
                    Width =10512
                    Height =7272
                    BackColor =0
                    Name ="Image2"
                    Picture ="shadowBorder"
                    GUID = Begin
                        0xfa13d7864e1c6744aeff170d95aa0a08
                    End

                    LayoutCachedWidth =10512
                    LayoutCachedHeight =7272
                    BackThemeColorIndex =-1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Prevenir que el usuario mueva la sombra
    DoCmd.SelectObject acForm, Me.Name, False
    
End Sub