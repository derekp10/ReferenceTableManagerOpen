VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IReferenceTableManagerConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'IReferenceTableManagerConfig
'By Derek Pierce (12/6/2016)
'Ver. 1.0 12/6/2016

Public Property Get RefTableCollections() As Collection
    
End Property

Public Property Let RefTableCollections(ByRef RefCollection As Collection)

End Property

Public Property Get RefTableDatabaseLocation() As String

End Property

Public Property Get RefTableDBConnection() As ADODB.Connection
    'Supply a configured Active X Data Object Connection Object.
    'GetDBCon
End Property
