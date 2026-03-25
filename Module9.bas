Attribute VB_Name = "Module9"
Function ConvertNumberToText(Number As Double, maincurrency As String, subcurrency As String)
Dim Array1(0 To 9) As String
Dim Array2(0 To 9) As String
Dim Array3(0 To 9) As String
Dim MyNumber As String
Dim GetNumber As String
Dim ReadNumber As String
Dim My100 As String
Dim My10 As String
Dim My1 As String
Dim My11 As String
Dim My12 As String
Dim GetText As String
Dim Billion As String
Dim Million As String
Dim Thousand As String
Dim Hundred As String
Dim Fraction As String
Dim MyAnd As String
Dim i As Integer
Dim ReMark As String


If Number > 999999999999.99 Then Exit Function
If Number < 0 Then
Number = Number * -1
ReMark = "”«·» "
End If

If Number = 0 Then
ConvertNumberToText = "’›—"
Exit Function
End If

MyAnd = " Ê"
Array1(0) = ""
Array1(1) = "„«∆…"
Array1(2) = "„«∆ «‰"
Array1(3) = "À·«À„«∆…"
Array1(4) = "√—»⁄„«∆…"
Array1(5) = "Œ„”„«∆…"
Array1(6) = "” „«∆…"
Array1(7) = "”»⁄„«∆…"
Array1(8) = "À„«‰„«∆…"
Array1(9) = " ”⁄„«∆…"

Array2(0) = ""
Array2(1) = " ⁄‘—"
Array2(2) = "⁄‘—Ê‰"
Array2(3) = "À·«ÀÊ‰"
Array2(4) = "√—»⁄Ê‰"
Array2(5) = "Œ„”Ê‰"
Array2(6) = "” Ê‰"
Array2(7) = "”»⁄Ê‰"
Array2(8) = "À„«‰Ê‰"
Array2(9) = " ”⁄Ê‰"

Array3(0) = ""
Array3(1) = "Ê«Õœ"
Array3(2) = "«À‰«‰"
Array3(3) = "À·«À…"
Array3(4) = "√—»⁄…"
Array3(5) = "Œ„”…"
Array3(6) = "” …"
Array3(7) = "”»⁄…"
Array3(8) = "À„«‰Ì…"
Array3(9) = " ”⁄…"

GetNumber = Format(Number, "000000000000.00")

i = 0
Do While i < 15

If i < 12 Then
MyNumber = Mid$(GetNumber, i + 1, 3)
Else
MyNumber = "0" + Mid$(GetNumber, i + 2, 2)
End If

If (Mid$(MyNumber, 1, 3)) > 0 Then

ReadNumber = Mid$(MyNumber, 1, 1)
My100 = Array1(ReadNumber)
ReadNumber = Mid$(MyNumber, 3, 1)
My1 = Array3(ReadNumber)
ReadNumber = Mid$(MyNumber, 2, 1)
My10 = Array2(ReadNumber)

If Mid$(MyNumber, 2, 2) = 11 Then My11 = "≈ÕœÏ ⁄‘—…"
If Mid$(MyNumber, 2, 2) = 12 Then My12 = "≈À‰Ï ⁄‘—…"
If Mid$(MyNumber, 2, 2) = 10 Then My10 = "⁄‘—…"


 
If ((Mid$(MyNumber, 1, 1)) > 0) And ((Mid$(MyNumber, 2, 2)) > 0) Then My100 = My100 + MyAnd
If ((Mid$(MyNumber, 3, 1)) > 0) And ((Mid$(MyNumber, 2, 1)) > 1) Then My1 = My1 + MyAnd

GetText = My100 + My1 + My10

If ((Mid$(MyNumber, 3, 1)) = 1) And ((Mid$(MyNumber, 2, 1)) = 1) Then
GetText = My100 + My11
If ((Mid$(MyNumber, 1, 1)) = 0) Then GetText = My11
End If

If ((Mid$(MyNumber, 3, 1)) = 2) And ((Mid$(MyNumber, 2, 1)) = 1) Then
GetText = My100 + My12
If ((Mid$(MyNumber, 1, 1)) = 0) Then GetText = My12
End If

If (i = 0) And (GetText <> "") Then
If ((Mid$(MyNumber, 1, 3)) > 10) Then
Billion = GetText + " „·Ì«—"
Else
Billion = GetText + " „·Ì«—« "
If ((Mid$(MyNumber, 1, 3)) = 2) Then Billion = " „·Ì«—"
If ((Mid$(MyNumber, 1, 3)) = 2) Then Billion = " „·Ì«—‰"
End If
End If

If (i = 3) And (GetText <> "") Then

If ((Mid$(MyNumber, 1, 3)) > 10) Then
Million = GetText + " „·ÌÊ‰"
Else
Million = GetText + " „·«ÌÌ‰"
If ((Mid$(MyNumber, 1, 3)) = 1) Then Million = " „·ÌÊ‰"
If ((Mid$(MyNumber, 1, 3)) = 2) Then Million = " „·ÌÊ‰«‰"
End If
End If

If (i = 6) And (GetText <> "") Then
If ((Mid$(MyNumber, 1, 3)) > 10) Then
Thousand = GetText + " √·›"
Else
Thousand = GetText + " √·«›"
If ((Mid$(MyNumber, 3, 1)) = 1) Then Thousand = " √·›"
If ((Mid$(MyNumber, 3, 1)) = 2) Then Thousand = " √·›«‰"
End If
End If

If (i = 9) And (GetText <> "") Then Hundred = GetText
If (i = 12) And (GetText <> "") Then Fraction = GetText
End If

i = i + 3
Loop

If (Billion <> "") Then
If (Million <> "") Or (Thousand <> "") Or (Hundred <> "") Then Billion = Billion + MyAnd
End If

If (Million <> "") Then
If (Thousand <> "") Or (Hundred <> "") Then Million = Million + MyAnd
End If

If (Thousand <> "") Then
If (Hundred <> "") Then Thousand = Thousand + MyAnd
End If

If Fraction <> "" Then
If (Billion <> "") Or (Million <> "") Or (Thousand <> "") Or (Hundred <> "") Then
ConvertNumberToText = ReMark + Billion + Million + Thousand + Hundred + " " + maincurrency + MyAnd + Fraction + " " + subcurrency
Else
ConvertNumberToText = ReMark + Fraction + " " + subcurrency
End If
Else
ConvertNumberToText = ReMark + Billion + Million + Thousand + Hundred + " " + maincurrency
End If
End Function
