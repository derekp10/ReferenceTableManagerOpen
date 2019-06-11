VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericReferenceTableConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'GenericReferenceTableConfig
'By Derek Pierce
'Ver. 1.0 6/10/2019

Implements IReferenceTableConfig

'Used to store the unique RefID, RefName, and RefTableName for a table in a database.


Private ref_ID As String
Private ref_Name As String
Private ref_TableName As String

Public Function SetConfigValues(ByVal RefTableName As String, Optional RefID As String = vbNullString, Optional RefName As String = vbNullString)
    
    ref_TableName = RefTableName
    
    If RefID <> vbNullString Then
        ref_ID = RefID
    Else:
        ref_ID = "RefID"
    End If
    
    If RefName <> vbNullString Then
        ref_Name = RefName
    Else:
        ref_Name = "RefName"
    End If
    
End Function


Private Property Get IReferenceTableConfig_RefID() As String
    IReferenceTableConfig_RefID = ref_ID
End Property

Private Property Get IReferenceTableConfig_RefName() As String
    IReferenceTableConfig_RefName = ref_Name
End Property

Private Property Get IReferenceTableConfig_RefTableName() As String
    IReferenceTableConfig_RefTableName = ref_TableName
End Property
