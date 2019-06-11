VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IReferenceTableConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'IReferenceTableConfig
'Ver. 1.0 6/10/2019

Public Property Get RefID() As String
    'Returns the field name that represents the unique ID for the table.
End Property

Public Property Get RefName() As String
    'Returns the field name that represents the unique name for the table.
End Property

Public Property Get RefTableName() As String
    'Returns the name of the table that RefID and RefName are attached to.
End Property
