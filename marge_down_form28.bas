Attribute VB_Name = "marge_down_form28"
Sub marge_down_form28_1()
Dim al_munatham As String
Dim al_muhaseb As String
Dim al_mudakek As String
Dim al_mudeerhesabat As String
Dim al_mudeermalia As String
Dim al_mudeermassnaa As String


al_munatham = ""
al_muhaseb = ""
al_mudakek = ""
al_mudeerhesabat = ""
al_mudeermalia = ""
al_mudeermassnaa = ""
UserForm28.TextBox4.value = Sheets(36).Cells(Rows.count, 1).End(xlUp).row + 4

'====================================================
Sheets(36).Range("a9  : " & "ce" & UserForm28.TextBox4.value + 3).UnMerge
'Range("a9  : " & "br" & UserForm28.TextBox4.Value + 3).

'Sheets(36).Range("a9  : " & "br" & UserForm28.TextBox4.Value + 3).
'=================================
Sheets(36).Range("a" & UserForm28.TextBox4.value + 3 & " : " & "ce" & UserForm28.TextBox4.value).Merge

'Sheets(36).Range("f" & UserForm28.TextBox4.Value + 3 & " : " & "g" & UserForm28.TextBox4.Value).Merge
'Sheets(36).Range("i" & UserForm28.TextBox4.Value + 3 & " : " & "k" & UserForm28.TextBox4.Value).Merge
'Sheets(36).Range("m" & UserForm28.TextBox4.Value + 3 & " : " & "o" & UserForm28.TextBox4.Value).Merge
'Sheets(36).Range("q" & UserForm28.TextBox4.Value + 3 & " : " & "s" & UserForm28.TextBox4.Value).Merge





al_munatham = " «·„‰Ÿ„"
al_muhaseb = "«·„Õ«”»"
al_mudakek = "«·„œﬁﬁ "
al_mudeerhesabat = "„œÌ— «· œﬁÌﬁ"
al_mudeermalia = "„œÌ— «·„«·Ì  "
al_mudeermassnaa = "„œÌ— «·„’‰⁄"
Sheets(36).Range("a" & UserForm28.TextBox4.value + 3 & " : " & "br" & UserForm28.TextBox4.value).value = al_munatham & "                            " & al_muhaseb _
& "                         " & al_mudakek & "                         " & al_mudeerhesabat _
& "                            " & al_mudeermassnaa
Sheets(36).Range("a" & UserForm28.TextBox4.value + 3 & " : " & "br" & UserForm28.TextBox4.value).HorizontalAlignment = xlCenter

End Sub
