Attribute VB_Name = "Module10"
Sub selectFile()
'On Error GoTo emad
'Create and set dialog box as variable
Dim dialogBox As FileDialog
Set dialogBox = Application.FileDialog(msoFileDialogOpen)

'Do not allow multiple files to be selected
dialogBox.AllowMultiSelect = False

'Set the title of the DialogBox
dialogBox.title = "Select a file"

'Set the default folder to open
dialogBox.InitialFileName = UserForm14.TextBox5.Text

'Clear the dialog box filters
dialogBox.Filters.Clear
'Apply file filters - use ; to separate filters for the same name
dialogBox.Filters.Add "image_document", "*.JPG;*.JPEG;*.PNG;*.TIFF;*.GIF;*.PDF;*.AI;*.RAW;*.xlsx;*.xls;*.xlsm"

'Show the dialog box and output full file name
If dialogBox.Show = -1 Then
  UserForm14.TextBox5.Text = dialogBox.SelectedItems(1)
  Dim last1, frw1 As Integer
last1 = Sheets(21).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(21).Cells(frw1, "b") = TextBox2.Text And Sheets(21).Cells(frw1, "c") = TextBox3.Text Then
Image1.Picture = LoadPicture(TextBox5.Text)

End If
Next
End If

'emad:
End Sub
