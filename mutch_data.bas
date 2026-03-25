Attribute VB_Name = "mutch_data"
Sub mutch_data_22()


Sheets(1).Activate
If UserForm48.TextBox2.Text = "" And UserForm48.TextBox5.Text = "" Then
UserForm48.Image1.Picture = LoadPicture("D:\employ_pic\shdow_pic.jpg")
End If
For i = 0 To UserForm48.ListBox1.ListCount
    If UserForm48.ListBox1.Selected(i) = True Then
        For j = 1 To 63
       UserForm48.Controls("userform48.TextBox" & j).Text = Cells(UserForm48.ListBox1.List(i, 1), j)
       UserForm48.TextBox10.Visible = False
        Next j
    End If
Next i
UserForm48.ComboBox1.Visible = False
'===================================
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "b") = UserForm48.TextBox2.Text And Sheets(1).Cells(frw1, "e") = UserForm48.TextBox5.Text Then
Image1.Picture = LoadPicture(Sheets(1).Cells(frw1, 64))
Image2.Picture = LoadPicture(Sheets(1).Cells(frw1, 65))
UserForm48.ComboBox4.Text = Sheets(1).Cells(frw1, 4)
UserForm48.ComboBox5.Text = Sheets(1).Cells(frw1, 6)
UserForm48.ComboBox6.Text = Sheets(1).Cells(frw1, "m")
UserForm48.ComboBox7.Text = Sheets(1).Cells(frw1, "o")
UserForm48.ComboBox8.Text = Sheets(1).Cells(frw1, "q")
UserForm48.ComboBox9.Text = Sheets(1).Cells(frw1, "ay")
UserForm48.ComboBox10.Text = Sheets(1).Cells(frw1, "z")
UserForm48.ComboBox11.Text = Sheets(1).Cells(frw1, "ab")
UserForm48.ComboBox3.Text = Sheets(1).Cells(frw1, "bf")

UserForm48.TextBox82.Text = Format(Sheets(1).Cells(frw1, "bn"), "yyyy/mm/dd")

UserForm48.TextBox83.Text = Sheets(1).Cells(frw1, "bo")
UserForm48.TextBox84.Text = Sheets(1).Cells(frw1, "bp")
UserForm48.TextBox85.Text = Sheets(1).Cells(frw1, "bq")
UserForm48.TextBox86.Text = Sheets(1).Cells(frw1, "br")
UserForm48.TextBox87.Text = Sheets(1).Cells(frw1, "bs")
UserForm48.TextBox88.Text = Sheets(1).Cells(frw1, "bt")
UserForm48.TextBox91.Text = Sheets(1).Cells(frw1, "bv")
UserForm48.TextBox99.Text = Sheets(1).Cells(frw1, "bu")
UserForm48.TextBox100.Text = Sheets(1).Cells(frw1, "ca")
UserForm48.TextBox105.Text = Sheets(1).Cells(frw1, "ce")

UserForm48.CheckBox4.value = Sheets(1).Cells(frw1, "cd")

If UserForm48.TextBox58.Text = "⁄ﬁœ" Then
UserForm48.TextBox10.Enabled = False
Label123.Visible = True
UserForm48.TextBox10.Text = Sheets(1).Cells(frw1, "j")
UserForm48.TextBox10.Visible = False
UserForm48.TextBox101.Visible = True

UserForm48.TextBox101.Text = Sheets(1).Cells(frw1, "j")

Else
UserForm48.TextBox101.Visible = False
UserForm48.TextBox10.Visible = True
Label123.Visible = False
End If
End If
Next frw1
If UserForm48.TextBox49.Text <> "" Then
UserForm48.TextBox77.Text = ConvertNumberToText(UserForm48.TextBox49.Text, "œÌ‰«—", "")
Else
UserForm48.TextBox77.Text = 0
End If
If UserForm48.TextBox7.Text = "„Ã«“Ì‰" Then
MsgBox "«·„ÊŸ› " & " " & UserForm48.TextBox5.Text & "  " & "«·„Êﬁ› «·ÊŸÌ›Ì ·Â „‰ «·„Ã«“Ì‰ Õ«Ê· «⁄«œ… «Õ ”«» «·„Êﬁ› «·„«·Ì ·Â ·Ì ÿ«»ﬁ „⁄ ﬁ«‰Ê‰ «·«Ã«“«  ", vbInformation + vbOKOnly, "—”«·…  ‰»ÌÂ"


End If
UserForm48.TextBox89.value = Sheets(1).Cells(2, "bw").value
UserForm48.TextBox90.value = Sheets(1).Cells(2, "bx").value

If UserForm48.TextBox58.Text = "—« » Ã“∆Ì" Or UserForm48.TextBox58.Text = "«„Ê„Â À«‰Ì 6 «‘Â—" Then
CommandButton16.Enabled = False
Else
UserForm48.CommandButton16.Enabled = True
End If

'===========================================
Sheets(1).Cells(4, "cb").value = UserForm48.TextBox58.Text

UserForm48.TextBox93.value = Sheets(1).Cells(4, "cd")
UserForm48.TextBox94.value = Sheets(1).Cells(4, "ce")
UserForm48.TextBox97.value = Sheets(1).Cells(4, "cf")
UserForm48.TextBox98.value = Sheets(1).Cells(4, "cg")
UserForm48.TextBox95.value = Sheets(1).Cells(4, "ch")
UserForm48.TextBox96.value = Sheets(1).Cells(4, "cc")
  
'userform48.TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

UserForm48.TextBox75.Text = UserForm48.ListView1.ListItems.count
UserForm48.TextBox93.value = Format(UserForm48.TextBox93.value, "#,## IQD")
UserForm48.TextBox94.value = Format(UserForm48.TextBox94.value, "#,## IQD")
UserForm48.TextBox97.value = Format(UserForm48.TextBox97.value, "#,## IQD")
UserForm48.TextBox98.value = Format(UserForm48.TextBox98.value, "#,## IQD")
UserForm48.TextBox95.value = Format(UserForm48.TextBox95.value, "#,## IQD")
UserForm48.TextBox96.value = Format(UserForm48.TextBox96.value, "#,## IQD")

UserForm48.Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & UserForm48.TextBox58.Text

UserForm48.TextBox89.value = Format(val(UserForm48.TextBox85) / val(UserForm48.TextBox10) * 100, "00.0") & "%"
UserForm48.TextBox90.value = Format(val(UserForm48.TextBox86) / val(UserForm48.TextBox10) * 100, "00.0") & "%"

UserForm48.ComboBox12.value = ""
UserForm48.ComboBox13.value = ""
UserForm48.TextBox103.value = "„ﬁœ«— ‰”»… «·„‹‹‹Œ’’«  «·„‹‹∆ÊÌ…" & " = " & (val(UserForm48.TextBox14) + val(UserForm48.TextBox16) + val(UserForm48.TextBox18) + val(UserForm48.TextBox27) + val(UserForm48.TextBox29)) / val(UserForm48.TextBox10) * 100 & "%"

End Sub
