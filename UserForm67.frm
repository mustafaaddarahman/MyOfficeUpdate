VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm67 
   Caption         =   "UserForm67"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8232.001
   OleObjectBlob   =   "UserForm67.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm67"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Call CalculateAG
Call FillAT
Call CalculateSalaryPositions
Call UpdateLiveFormulas_Sheet1
Call UpdateSummaryFormula
Call CalculateGradesAndStatus
Call count_all
Call SaveWorkbook
End Sub

Private Sub CommandButton2_Click()

Call CalculateFiveYearLeaveOnly
Call ClearAG
Call FillAT
Call CalculateSalaryPositions
Call UpdateLiveFormulas_Sheet1
Call UpdateSummaryFormula
Call CalculateFiveYearLeaveOnly
Call CalculateGradesAndStatus
Call count_all

Call SaveWorkbook
End Sub

Private Sub CommandButton3_Click()
Unload Me

End Sub

Private Sub CommandButton4_Click()

End Sub
