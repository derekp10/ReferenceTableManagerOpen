Attribute VB_Name = "Constants"
Option Explicit

Public Const DEBUG_MODE = True

Public Function DEV_EXPORT_LOCATION() As String
    If Not DEBUG_MODE Then
        'Do nothing test only
    Else:
        DEV_EXPORT_LOCATION = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1) & "\src\"
    End If
End Function
