Option Compare Database
Option Explicit

' ----------------------------------------------------------------
' Procedure : CreateModernOnOff
' DateTime  : 7/26/2022 22:57
' Author    : Mike Wolfe
' Source    : https://nolongerset.com/modern-on-off/
' Purpose   : Create a modern-looking on/off button on an Access form.
' ----------------------------------------------------------------
Sub CreateModernOnOff(FormName As String, _
                      Optional IncludeOnOffTextBox As Boolean = False)
    'Open the form in design view...
    If SysCmd(acSysCmdGetObjectState, acForm, FormName) = 0 Then
        '...if it's not already open or...
        DoCmd.OpenForm FormName, acDesign
    ElseIf Forms(FormName).CurrentView <> 0 Then
        '...if it is open but not in Design view
        DoCmd.OpenForm FormName, acDesign
    End If
    
    'Create the toggle button in the upper left corner of the form
    '   (you can move it manually in the form designer later)
    Dim ToggleBtn As ToggleButton
    Set ToggleBtn = CreateControl(FormName, acToggleButton)
    With ToggleBtn
        'Disable bevel effect
        .Bevel = 0
        
        'Ensure shape is a rounded rectangle
        .Shape = 2
        
        'Disable all tinting and shading
        .BackTint = 100
        .BorderTint = 100
        .HoverTint = 100
        .PressedShade = 100
    
        'Set border properties
        .BorderStyle = 1 'Solid
        .BorderWidth = 3
        
        'Set optimal height/width
        .Height = 300
        .Width = 780
        
        'Set Font properties
        .FontName = "Wingdings"
        .FontSize = 10
        .FontBold = True
        
        'Set properties for when the control is OFF
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .HoverColor = vbWhite
        
        'Set properties for when the control is ON
        .PressedColor = vbBlue
        .PressedForeColor = vbWhite
        
        'Call the FormatToggle() function whenever the control is toggled ON/OFF
        .AfterUpdate = "=FormatToggle([ActiveControl])"
        
        'Format the control as OFF for design purposes
        .Caption = "l  "   'lowercase "l" is a filled circle glyph in the Wingdings font
        .BorderColor = vbBlack

    End With
    
    Debug.Print "Be sure to add the line..." & vbNewLine
    Debug.Print "    FormatToggle Me." & ToggleBtn.Name & vbNewLine
    Debug.Print "...to the Form_Current() event handler for form " & FormName
    
    If IncludeOnOffTextBox Then
        Dim TxtBox As TextBox
        Set TxtBox = CreateControl(FormName, acTextBox)
        With TxtBox
            .ControlSource = "=IIf(" & ToggleBtn.Name & ", 'ON', 'OFF')"
            .BorderStyle = 0 'Transparent
            .Enabled = False
            .Locked = True
            .Left = ToggleBtn.Width + 100
            .Height = ToggleBtn.Height
        End With
        
    End If
End Sub


Function FormatToggle(ToggleBtn As ToggleButton)
    Const Buffer As Long = 2
    With ToggleBtn
        If .Value Then
            'When button is ON, leading spaces force
            '   disc icon to the right
            .Caption = Space(Buffer) & "l"
            .BorderColor = vbBlue
        Else
            'When button is OFF, trailing spaces force
            '   disc icon to the left
            .Caption = "l" & Space(Buffer)
            .BorderColor = vbBlack
        End If
    End With
End Function