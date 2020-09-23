Attribute VB_Name = "modMain"
Sub main()
    'check if the splash screen should be shown or not
    showsplash = CLng(GetSetting("mousewrap", "splash", "enabled", 1))
    If showsplash <> 0 Then
        'show the splashscreen
        frm_main.Show
    Else
        'move the form off the screen so the user wont see it
        frm_main.Left = -frm_main.Width - 10
    End If
End Sub
