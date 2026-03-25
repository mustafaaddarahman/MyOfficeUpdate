VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm71 
   Caption         =   "UserForm71"
   ClientHeight    =   11616
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18768
   OleObjectBlob   =   "UserForm71.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm71"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Dim WithEvents TextBox1 As MSForms.TextBox
Attribute TextBox1.VB_VarHelpID = -1
Dim WithEvents TextBox4 As MSForms.TextBox
Attribute TextBox4.VB_VarHelpID = -1
Dim WithEvents TextBox6 As MSForms.TextBox
Attribute TextBox6.VB_VarHelpID = -1
Dim WithEvents ListBox1 As MSForms.ListBox
Attribute ListBox1.VB_VarHelpID = -1
Dim WithEvents ListBox2 As MSForms.ListBox
Attribute ListBox2.VB_VarHelpID = -1
Dim WithEvents CommandButton1 As MSForms.CommandButton
Attribute CommandButton1.VB_VarHelpID = -1
Dim WithEvents CommandButton2 As MSForms.CommandButton
Attribute CommandButton2.VB_VarHelpID = -1
Dim WithEvents CommandButton3 As MSForms.CommandButton
Attribute CommandButton3.VB_VarHelpID = -1
Dim WithEvents CommandButton4 As MSForms.CommandButton
Attribute CommandButton4.VB_VarHelpID = -1
Dim WithEvents CommandButton5 As MSForms.CommandButton ' “— Õ–› „‰ «·ﬁ«∆„… «·„ƒﬁ …
Attribute CommandButton5.VB_VarHelpID = -1

Dim TempData As New Collection

Private Sub CommandButton3_Click()

End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim lbls As Variant
    Dim i As Integer
    
    ' ÷»ÿ ÕÃ„ «·›Ê—„ «·ﬂ·Ì
    Me.Width = 950
    Me.Height = 720

    ' 1. »‰«¡  Ìﬂ”  »Êﬂ” «·»ÕÀ
    Set TextBox1 = Me.Controls.Add("Forms.TextBox.1", "TextBox1")
    With TextBox1
        .Left = 15
        .Top = 15
        .Width = 220
        .Height = 30
        ApplyStyle TextBox1
    End With
    
    ' 2. »‰«¡ ﬁ«∆„… «·„ÊŸ›Ì‰ «·—∆Ì”Ì…
    Set ListBox1 = Me.Controls.Add("Forms.ListBox.1", "ListBox1")
    With ListBox1
        .Left = 15
        .Top = 55
        .Width = 220
        .Height = 300
        .ColumnCount = 2
        .ColumnWidths = "50;150"
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With

    ' 3. »‰«¡ ‘«‘… ⁄—÷ «·„÷«›Ì‰ (ListBox2)
    With Me.Controls.Add("Forms.Label.1")
        .Caption = "«·„ÊŸ›Ì‰ «·„÷«›Ì‰ ··„Ã„Ê⁄… (ﬁ»· «· —ÕÌ·):"
        .Left = 15
        .Top = 365
        .Width = 300
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With

    Set ListBox2 = Me.Controls.Add("Forms.ListBox.1", "ListBox2")
    With ListBox2
        .Left = 15
        .Top = 385
        .Width = 620
        .Height = 250
        .ColumnCount = 8
        .ColumnWidths = "50;120;50;50;80;80;60;80"
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With

    ' 4. „’›Ê›… «·⁄‰«ÊÌ‰ ÃÂ… «·Ì„Ì‰
    lbls = Array("«·—ﬁ„ «·ÊŸÌ›Ì", "«·œ—Ã… «·”«»ﬁ…", "«·œ—Ã… «·Õ«·Ì…", "«·„—Õ·… «·”«»ﬁ…", _
                 "«·„—Õ·… «·Õ«·Ì…", "«·«” Õﬁ«ﬁ «·”«»ﬁ", "«·«” Õﬁ«ﬁ «·Õ«·Ì", "«·«”„Ì «·”«»ﬁ („’ÕÕ)", _
                 "«·«”„Ì «·Õ«·Ì", "«·—« » «·ﬂ·Ì «·”«»ﬁ", "«·—« » «·ﬂ·Ì «·Õ«·Ì", "„ﬁœ«— ⁄·«Ê… «·«”„Ì", "›—ﬁ «·—« » «·ﬂ·Ì (⁄—÷)")
    
    ' 5. »‰«¡ «··Ì»·«  Ê«· Ìﬂ”  »Êﬂ”
    For i = 0 To UBound(lbls)
        With Me.Controls.Add("Forms.Label.1")
            .Caption = lbls(i)
            .Left = 650
            .Top = 15 + (i * 35)
            .Width = 140
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .Font.Bold = True
        End With
        
        With Me.Controls.Add("Forms.TextBox.1", "TextBox" & (i + 2))
            .Left = 800
            .Top = 15 + (i * 35)
            .Width = 120
            .Height = 28
            ApplyStyle Me.Controls("TextBox" & (i + 2))
            If i = 12 Then
                .Locked = True
                .BackColor = &HE0E0E0
            End If
        End With
    Next i

    ' —»ÿ «·„ €Ì—«  »«·√Õœ«À
    Set TextBox4 = Me.Controls("TextBox4")
    Set TextBox6 = Me.Controls("TextBox6")
    
    ' 6. »‰«¡ «·√“—«—
    Set CommandButton1 = Me.Controls.Add("Forms.CommandButton.1", "CommandButton1")
    With CommandButton1
        .Caption = "≈÷«›… «·„ÊŸ› ··„Ã„Ê⁄…"
        .Left = 650
        .Top = 480
        .Width = 270
        .Height = 35
        .Font.Bold = True
        .BackColor = &HFFC0C0
    End With

    Set CommandButton5 = Me.Controls.Add("Forms.CommandButton.1", "CommandButton5")
    With CommandButton5
        .Caption = "Õ–› «·„Œ «— „‰ «·ﬁ«∆„… «·„ƒﬁ …"
        .Left = 650
        .Top = 520
        .Width = 270
        .Height = 35
        .Font.Bold = True
        .BackColor = &HC0C0FF
    End With

    Set CommandButton2 = Me.Controls.Add("Forms.CommandButton.1", "CommandButton2")
    With CommandButton2
        .Caption = " —ÕÌ· Ê ÕœÌÀ ‘«„· "
        .Left = 650
        .Top = 560
        .Width = 270
        .Height = 35
        .Font.Bold = True
        .BackColor = &H80FF80
    End With

    Set CommandButton4 = Me.Controls.Add("Forms.CommandButton.1", "CommandButton4")
    With CommandButton4
        .Caption = "≈€·«ﬁ"
        .Left = 650
        .Top = 600
        .Width = 270
        .Height = 35
        .Font.Bold = True
        .BackColor = &H80C0FF
    End With

    UpdateList ""
End Sub

' --- “— 1: «·≈÷«›… ··„Ã„Ê⁄… («·–«ﬂ—… + «·ﬁ«∆„… «·„ƒﬁ …) ---
Private Sub CommandButton1_Click()
    Dim arr(1 To 9) As Variant
    Dim EmpName As String
    
    ' «· Õﬁﬁ „‰ «Œ Ì«— „ÊŸ›
    If ListBox1.ListIndex = -1 Then
        MsgBox "Ì—ÃÏ «Œ Ì«— „ÊŸ› „‰ «·ﬁ«∆„… √Ê·«"
        Exit Sub
    End If
    
    EmpName = ListBox1.List(ListBox1.ListIndex, 1)
    
    '  ÃÂÌ“ «·»Ì«‰«  ›Ì «·„’›Ê›…
    arr(1) = ListBox1.value ' «·—ﬁ„ «·ÊŸÌ›Ì
    arr(2) = EmpName ' «·«”„
    arr(3) = TextBox4.value ' «·œ—Ã… «·ÃœÌœ…
    arr(4) = TextBox6.value ' «·„—Õ·… «·ÃœÌœ…
    arr(5) = val(Me.Controls("TextBox7").value) ' «·«” Õﬁ«ﬁ «·”«»ﬁ
    arr(6) = val(Me.Controls("TextBox8").value) ' «·«” Õﬁ«ﬁ «·Õ«·Ì
    arr(7) = val(Me.Controls("TextBox14").value) ' «·›—ﬁ
    arr(8) = val(Me.Controls("TextBox12").value) ' «·’«›Ì «·Õ«·Ì
    arr(9) = val(Me.Controls("TextBox10").value) ' «·«”„Ì «·ÃœÌœ ·· ÕœÌÀ
    
    ' ≈÷«›… ··„’›Ê›… («·–«ﬂ—…)
    TempData.Add arr
    
    ' «· ⁄«„· „⁄ ListBox2
    With ListBox2
        ' --- ≈⁄œ«œ«  «· ‰”Ìﬁ Ê«·‘ﬂ· ---
        .ColumnCount = 8
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
        
        .TextAlign = fmTextAlignCenter
        
        ' ÷»ÿ ⁄—÷ «·√⁄„œ… (»«·‰ﬁ«ÿ) - Ì„ﬂ‰ﬂ  ⁄œÌ· «·√—ﬁ«„ · ‰«”» ÕÃ„ «·‘«‘… ·œÌﬂ
        ' «·—ﬁ„ «·ÊŸÌ›Ì° «·«”„° «·œ—Ã…° «·„—Õ·…° «·”«»ﬁ° «·Õ«·Ì° «·›—ﬁ° «·’«›Ì
        .ColumnWidths = "60;120;50;50;80;80;60;80"
        
        ' --- ≈÷«›… «·⁄‰«ÊÌ‰ ﬂ√Ê· ’› ≈–« ﬂ«‰  «·ﬁ«∆„… ›«—€… ---
        If .ListCount = 0 Then
            .AddItem "«·—ﬁ„"
            .List(0, 1) = "«”„ «·„ÊŸ›"
            .List(0, 2) = "«·œ—Ã…"
            .List(0, 3) = "«·„—Õ·…"
            .List(0, 4) = "«·«” Õﬁ«ﬁ ”"
            .List(0, 5) = "«·«” Õﬁ«ﬁ Õ"
            .List(0, 6) = "«·›—ﬁ"
            .List(0, 7) = "«·’«›Ì"
        End If
        
        ' --- ≈÷«›… «·»Ì«‰«  «·›⁄·Ì… ---
        .AddItem arr(1)
        .List(.ListCount - 1, 1) = arr(2)
        .List(.ListCount - 1, 2) = arr(3)
        .List(.ListCount - 1, 3) = arr(4)
        .List(.ListCount - 1, 4) = Format(arr(5), "#,##0")
        .List(.ListCount - 1, 5) = Format(arr(6), "#,##0")
        .List(.ListCount - 1, 6) = Format(arr(7), "#,##0")
        .List(.ListCount - 1, 7) = Format(arr(8), "#,##0")
        
        ' «· „—Ì— ·√”›·  ·ﬁ«∆Ì« ·—ƒÌ… ¬Œ— ≈÷«›…
        .ListIndex = .ListCount - 1
    End With
End Sub

' --- “— 5: Õ–› ﬁÌœ „‰ «·ﬁ«∆„… «·„ƒﬁ … ﬁ»· «· —ÕÌ· ---
Private Sub CommandButton5_Click()
On Error Resume Next
    If ListBox2.ListIndex <> -1 Then
        TempData.Remove (ListBox2.ListIndex + 1)
        ListBox2.RemoveItem (ListBox2.ListIndex)
    End If
End Sub

' --- “— 2: «· —ÕÌ· «·›⁄·Ì Ê«· ÕœÌÀ «·‘«„· ·‹ 1 Ê 35 ---
Private Sub CommandButton2_Click()
    Dim Item As Variant
    Dim r As Range
    Dim tRow As Long
    Dim sRow As Long
    
    If TempData.count = 0 Then
        MsgBox "«·ﬁ«∆„… «·„ƒﬁ … ›«—€…"
        Exit Sub
    End If
    
    For Each Item In TempData
        ' «·»ÕÀ ›Ì ‘Ì  1
        Set r = Sheets(1).Columns("B").Find(Item(1), LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not r Is Nothing Then
            sRow = r.row
            tRow = Sheets(35).Cells(Sheets(35).Rows.count, "B").End(xlUp).row + 1
            If tRow < 9 Then
                tRow = 9
            End If
            
            ' √- «· —ÕÌ· ·‹ 35
            Sheets(1).Range("A" & sRow & ":BU" & sRow).Copy
            Sheets(35).Range("A" & tRow).PasteSpecial xlPasteValues
            
            With Sheets(35)
                .Cells(tRow, "BV").value = Item(6) - Item(5) ' «·›—ﬁ
                .Cells(tRow, "BW").value = Item(3) ' «·œ—Ã…
                .Cells(tRow, "BX").value = Item(4) ' «·„—Õ·…
                .Cells(tRow, "BY").value = Item(9) ' «·«”„Ì «·ÃœÌœ
                .Cells(tRow, "BZ").value = Item(6) ' «·«” Õﬁ«ﬁ
                .Cells(tRow, "CB").value = Item(8) ' «·ﬂ·Ì
                .Cells(tRow, "CD").value = Format(Date, "mmmm", vbArabic)
                .Cells(tRow, "CE").value = "«÷«›Â"
                .Cells(tRow, "CF").value = "„ €Ì—«  «÷«›… «·⁄·«Ê« "
            End With
            
            ' »- «· ÕœÌÀ «·‘«„· ·‹ 1
            Sheets(1).Cells(sRow, "J").value = Item(9)
            Sheets(1).Cells(sRow, "H").value = Item(3)
            Sheets(1).Cells(sRow, "I").value = Item(4)
            Sheets(1).Cells(sRow, "BZ").value = Item(6)
            Sheets(1).Cells(sRow, "CB").value = Item(8)
            Sheets(1).Cells(sRow, "BV").value = Item(6) - Item(5)
        End If
    Next Item
    
    '  ›—Ì€ «·–«ﬂ—… Ê«·‘«‘…
    Set TempData = New Collection
    ListBox2.Clear
    MsgBox " „  —ÕÌ· ﬂ«›… «·ﬁÌÊœ Ê ÕœÌÀ «·”Ã· «·—∆Ì”Ì »‰Ã«Õ"

Sheets(34).Activate
Sheets(35).Activate

End Sub

' --- «·„Õ—ﬂ «·„«·Ì Ê«·›· —… ---

Private Sub CalculateFinances()
    On Error Resume Next
    Dim r As Range
    Dim sRow As Long
    Dim OldN As Double, NewN As Double
    Dim EstH As Double, NetH As Double, ExtraBonus As Double
    Dim NewDegree As Variant
    
    ' 1. Ã·» «·—ﬁ„ «·ÊŸÌ›Ì „‰ TextBox2 ··Ê’Ê· ··ﬁÌœ
    Dim EmpID As String: EmpID = Me.Controls("TextBox2").value
    If EmpID = "" Then Exit Sub
    
    '  ’ÕÌÕ «·”‰ ﬂ” Â‰«: ≈÷«›… «·‰ﬁÿ… ﬁ»· xlWhole
    Set r = Sheets(1).Columns("B").Find(EmpID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If r Is Nothing Or r.row < 9 Then Exit Sub
    sRow = r.row
    
    ' --- √Ê·«: Ã·» «·»Ì«‰«  «·”«»ﬁ… «·„”Ã·… ›Ì «·ﬁÌœ ---
    Dim OldNetS As Double: OldNetS = val(Sheets(1).Cells(sRow, "AW").value)
    Me.Controls("TextBox9").value = val(Sheets(1).Cells(sRow, "J").value)
    Me.Controls("TextBox7").value = val(Sheets(1).Cells(sRow, "AE").value)
    Me.Controls("TextBox11").value = OldNetS

    ' --- À«‰Ì«: «” Œ—«Ã «·—« » «·«”„Ì «·ÃœÌœ „‰ ‘Ì  9 ---
    Sheets(9).Range("G18").value = Me.Controls("TextBox4").value
    Sheets(9).Range("F18").value = Me.Controls("TextBox6").value
    NewN = val(Sheets(9).Range("H18").value)
    Me.Controls("TextBox10").value = NewN
    
    ' --- À«·À«: Ã·»  ›«’Ì· „Œ’’«  «·ﬁÌœ (sRow) ---
    Dim m As Double: m = val(Sheets(1).Cells(sRow, "M").value)
    Dim o As Double: o = val(Sheets(1).Cells(sRow, "O").value)
    Dim q As Double: q = val(Sheets(1).Cells(sRow, "Q").value)
    Dim z As Double: z = val(Sheets(1).Cells(sRow, "Z").value)
    Dim ab As Double: ab = val(Sheets(1).Cells(sRow, "AB").value)
    
    ' «·„Œ’’«  «·„ﬁÿÊ⁄… Ê«·«” ﬁÿ«⁄« 
    Dim s As Double: s = val(Sheets(1).Cells(sRow, "S").value)
    Dim t As Double: t = val(Sheets(1).Cells(sRow, "T").value)
    Dim u As Double: u = val(Sheets(1).Cells(sRow, "U").value)
    Dim Y As Double: Y = val(Sheets(1).Cells(sRow, "Y").value)
    Dim bh As Double: bh = val(Sheets(1).Cells(sRow, "BH").value)
    Dim bp As Double: bp = val(Sheets(1).Cells(sRow, "BP").value)
    Dim br As Double: br = val(Sheets(1).Cells(sRow, "BR").value)
    Dim av As Double: av = val(Sheets(1).Cells(sRow, "AV").value)

    ' --- —«»⁄«: ‘—ÿ «·œ—Ã«  (8° 9° 10) »‰«¡ ⁄·Ï «·œ—Ã… «·ÃœÌœ… ›Ì TextBox4 ---
    NewDegree = Me.Controls("TextBox4").value
    ExtraBonus = 0
    
    If NewDegree = "8" Or NewDegree = "9" Or NewDegree = "10" Or _
       val(NewDegree) = 8 Or val(NewDegree) = 9 Or val(NewDegree) = 10 Then
        ExtraBonus = NewN * 0.5
    End If

    ' --- Œ«„”«:  ÿ»Ìﬁ «·„⁄«œ·… «·‰Â«∆Ì… ··«” Õﬁ«ﬁ «·Õ«·Ì ---
    ' «” Õﬁ«ﬁ = «”„Ì ÃœÌœ + («”„Ì ÃœÌœ * „Ã„Ê⁄ «·‰”») + „Œ’’«  „ﬁÿÊ⁄… + «·“Ì«œ… «·Œ«’…
    EstH = NewN + (NewN * (m + o + q + z + ab)) + s + t + u + Y + bh + bp + br + ExtraBonus
    
    ' «·—« » «·ﬂ·Ì «·Õ«·Ì = «·«” Õﬁ«ﬁ «·Õ«·Ì - «·«” ﬁÿ«⁄« 
    NetH = EstH - av
    
    '  ÕœÌÀ «·›Ê—„ »«·‰ «∆Ã
    Me.Controls("TextBox8").value = EstH
    Me.Controls("TextBox12").value = NetH
    Me.Controls("TextBox13").value = NewN - val(Me.Controls("TextBox9").value)
    Me.Controls("TextBox14").value = NetH - OldNetS
    
End Sub
' --- «·√ﬂÊ«œ «·„”«⁄œ… «·„‰”ﬁ… ---
Private Sub TextBox1_Change(): UpdateList TextBox1.Text: End Sub
Private Sub TextBox4_Change(): CalculateFinances: End Sub
Private Sub TextBox6_Change(): CalculateFinances: End Sub
Private Sub CommandButton4_Click()
 
 Unload Me
 
End Sub


Private Sub ListBox1_Click()
    Dim r As Range
    Set r = Sheets(1).Columns("B").Find(ListBox1.value, LookIn:=xlValues, LookAt:=xlWhole)
    If Not r Is Nothing Then
        Me.Controls("TextBox2").value = Sheets(1).Cells(r.row, "B").value
        Me.Controls("TextBox3").value = Sheets(1).Cells(r.row, "H").value
        Me.Controls("TextBox5").value = Sheets(1).Cells(r.row, "I").value
        TextBox4.value = Sheets(1).Cells(r.row, "H").value
        TextBox6.value = val(Sheets(1).Cells(r.row, "I").value) + 1
        CalculateFinances
    End If
End Sub

Sub ApplyStyle(ctrl As Object)
    With ctrl
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
        .TextAlign = 2
    End With
End Sub

Sub UpdateList(SearchTxt As String)
    Dim i As Long
    Dim lastR As Long
    ListBox1.Clear
    lastR = Sheets(1).Cells(Sheets(1).Rows.count, "B").End(xlUp).row
    For i = 9 To lastR
        If InStr(1, Sheets(1).Cells(i, "E").value, SearchTxt, vbTextCompare) > 0 Or SearchTxt = "" Then
            ListBox1.AddItem Sheets(1).Cells(i, "B").value
            ListBox1.List(ListBox1.ListCount - 1, 1) = Sheets(1).Cells(i, "E").value
        End If
    Next i
End Sub


