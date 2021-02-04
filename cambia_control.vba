In an Excel 5/95 dialogsheet it is possible to change the value/content of a collection of controls by
looping through the controls in the collection, e.g. like this: For Each cb In dlg.CheckBoxes.
In Excel 97 or later the UserForm-object doesn't group the controls in the same way.
Below you will find some example macros that shows how
you can change the value/content of several UserForm-controls:

Sub ResetAllCheckBoxesInUserForm()
Dim ctrl As Control
    For Each ctrl In UserForm1.Controls
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.Value = False
        End If
    Next ctrl
End Sub

Sub ResetAllOptionButtonsInUserForm()
Dim ctrl As Control
    For Each ctrl In UserForm1.Controls
        If TypeName(ctrl) = "OptionButton" Then
            ctrl.Value = False
        End If
    Next ctrl
End Sub

Sub ResetAllTextBoxesInUserForm()
Dim ctrl As Control
    For Each ctrl In UserForm1.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Text = ""
        End If
    Next ctrl
End Sub
