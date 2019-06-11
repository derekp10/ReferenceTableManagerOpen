VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceTableManagerCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ReferenceTableManager
'By Derek Pierce (8/11/2015)
'Ver. 1.0 8/11/2015 First Reusable Build
'Ver. 1.1 3/28/2015
'Ver. 1.2 12/6/2016
'Ver. 1.2.1 3/24/2017 Bug Fix
'Ver. 1.3 6/10/2019

'Table rows are stored as RefTableDataClass
'All of the collections are built out of/contain this data class
Private RefTableData As RefTableDataClass

Private RefTableBuildDate As Collection
Private LastDBCheckTime As Date
Private blnIgnoreRefUpdateChecks As Boolean
Private Const RTM_EMPTY_DATE As Date = "12:00:00 AM"
Private RefTableConfig As IReferenceTableManagerConfig

'Builds the various RefTable collections base on the named table.
Private Sub Class_Initialize()
    Set RefTableBuildDate = New Collection
End Sub

Public Property Let SetRefTableConfig(ByRef Config As IReferenceTableManagerConfig)
    Set RefTableConfig = Config
End Property

Public Function CORE_GetRefCol(ByRef TargetRefTable As IReferenceTableConfig, Optional ByVal KeyType As RefTableKeyType = rtkt_RefTypeRefName) As Collection
    'Returns the local stored version of a RefTable as a collection for use with an optional RefTableKeyType Paramater if
    'the perefered collection key should be the RefID, or the RefName (Defaults to RefName)
    Dim rtnRefTblCol As RefTableCollectionClass
    
    If blnIgnoreRefUpdateChecks = False Then
        CORE_UpdateRefTableCollections , TargetRefTable
    End If

    Set rtnRefTblCol = RefTableConfig.RefTableCollections.Item(TargetRefTable.RefTableName)
    
    Set CORE_GetRefCol = rtnRefTblCol.GetCollectionByRefType(KeyType)
End Function

Private Function SetRefTableCollection(ByRef TargetRefTable As IReferenceTableConfig, ByRef NewRefTableCollectionData As RefTableCollectionClass)
    'Used by the CORE_UpdateRefTableCollections to set the specific RefTableCollectionClass data to a new
    '(if updated or forced) set of data from the database.
    Dim tmpRefTableCollection As RefTableCollectionClass

    If Not RTMInCollection(RefTableConfig.RefTableCollections, TargetRefTable.RefTableName) Then
        RefTableConfig.RefTableCollections.Add NewRefTableCollectionData, TargetRefTable.RefTableName
    Else:
        RefTableConfig.RefTableCollections.Remove TargetRefTable.RefTableName 'This change may not work needs testing.
        RefTableConfig.RefTableCollections.Add NewRefTableCollectionData, TargetRefTable.RefTableName
    End If
    
End Function

Public Function CORE_UpdateRefTableCollections(Optional ByVal ForceUpdate As Boolean = False, Optional ByRef TargetRefTable As IReferenceTableConfig)
    Dim tmpLastTime As Date
    Dim RefEnuNum As Long

    
'    If DateDiff("s", LastDBCheckTime, DateTime.Now) > 5 Then
'        tmpLastTime = LastRefTableUpdate
'        LastDBCheckTime = DateTime.Now
'    End If
    

    If Not RTMInCollection(RefTableBuildDate, TargetRefTable.RefTableName) Then
        tmpLastTime = LastRefTableUpdate(TargetRefTable)
        RefTableBuildDate.Add tmpLastTime, TargetRefTable.RefTableName
    End If
    

    If ForceUpdate = True Then
        tmpLastTime = LastRefTableUpdate(TargetRefTable)
        SetRefTableCollection TargetRefTable, BuildRefTableCollectionClass(TargetRefTable)
        RefTableBuildDate.Remove (TargetRefTable.RefTableName)
        RefTableBuildDate.Add tmpLastTime, TargetRefTable.RefTableName
    Else:
        If RefTableBuildDate.Item(TargetRefTable.RefTableName) < LastRefTableUpdate(TargetRefTable) Then
            tmpLastTime = LastRefTableUpdate(TargetRefTable)
            SetRefTableCollection TargetRefTable, BuildRefTableCollectionClass(TargetRefTable)
            RefTableBuildDate.Remove (TargetRefTable.RefTableName)
            RefTableBuildDate.Add tmpLastTime, TargetRefTable.RefTableName
        End If
    End If


    
End Function

'Using a string value and a reference table to look up, returns the ID for that string
Public Function CORE_GetIDFromString(ByRef TargetTypeRefTable As IReferenceTableConfig, ByVal strTarget As String) As Long
    'Key lookups are CASE SENSITIVE!!
    CORE_GetIDFromString = GetRelevantDataByStr(TargetTypeRefTable, strTarget).RefTypeID

End Function

'Using a long(variant) value and a reference table to look up, returns the String value for that ID
Public Function CORE_GetStringFromID(ByRef TargetTypeRefTable As IReferenceTableConfig, ByVal lngTarget As Variant) As String

    CORE_GetStringFromID = GetRelevantDataByID(TargetTypeRefTable, lngTarget).RefTypeName

End Function

'Gets any extra data that may be present for the Reftable. (Forward comments etc..)
Public Function CORE_GetExtraDataFromID(ByRef TargetTypeRefTable As IReferenceTableConfig, ByVal lngTarget As Variant) As RefTableExtraDataCollection
    
    Set CORE_GetExtraDataFromID = GetRelevantDataByID(TargetTypeRefTable, lngTarget).RefTypeExtraData
    
End Function

'Returns a RefTableDataClass for use with the GetXDatafromY functions based off a string search.
Private Function GetRelevantDataByStr(ByRef TargetTypeRefTable As IReferenceTableConfig, ByVal varTarget As String) As RefTableDataClass
    Dim tmpRefTableData As RefTableDataClass
    
    If varTarget <> "No Data" Then
        Set tmpRefTableData = CORE_GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefName).Item(CStr(varTarget))
    End If
    
    If tmpRefTableData Is Nothing Then
        Set tmpRefTableData = New RefTableDataClass
        tmpRefTableData.RefTypeID = 0
        tmpRefTableData.RefTypeName = ""
        tmpRefTableData.RefTypeExtraData = Nothing
    End If
    
    Set GetRelevantDataByStr = tmpRefTableData

End Function

'Returns a RefTableDataClass for use with the GetXDataFromY functions based off a long(variant) search.
Private Function GetRelevantDataByID(ByRef TargetTypeRefTable As IReferenceTableConfig, varTarget As Variant) As RefTableDataClass
    Dim tmpRefTableData As RefTableDataClass
    
    If RTMInCollection(CORE_GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefID), CStr(varTarget)) Then
        Set tmpRefTableData = CORE_GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefID).Item(CStr(varTarget))
    End If
    
'Old Blind to 0 Key code
'    If varTarget <> 0 Then
'        Set tmpRefTableData = GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefID).Item(CStr(varTarget))
'    End If
    
    If tmpRefTableData Is Nothing Then
        Set tmpRefTableData = New RefTableDataClass
        tmpRefTableData.RefTypeID = 0
        tmpRefTableData.RefTypeName = ""
        tmpRefTableData.RefTypeExtraData = Nothing
    End If
    
    Set GetRelevantDataByID = tmpRefTableData

End Function

'Class function used to create the various reference collections from the database data.
'Select TargetRefTable using the RefTableEnu enumerator, (If not set up it won't work.) get collection representing that tables data.
'Reqired Fields are RefID and RefName, and are used by the system to get the basic information required for this system to work.
'(A Query/View where the Primary Unique Identifier is aliased as RefID and a easy readable name as RefName can be used instead)
'NOTE: Above line is incorrect. ApplicationDataStore has no way of knowing about linkage of query to table and updates so
'       updates to primary table will not cause a change in ApplicationDataStore for the query entry.
'       Only use a Query or View if you know what you are doing. TODO: Fix this issue.
'Any fields not named RefID and RefName are added to a RefTableExtraDataCollection object inside of the RefTableDataClass.
'The Field.Name is used as the Key (Cast as String) for the collection, and the Field.Value is used as the Value.
'Not tested where the Field.Value is an object.
Private Function BuildRefTableCollectionClass(ByRef TargetRefTable As IReferenceTableConfig, Optional KeyType As RefTableKeyType = rtkt_RefTypeRefName) As RefTableCollectionClass
    'Generic function to build a collection based of a 2 field table. (IE: RefTable Type Tables)
    Dim strSQL As String
    Dim rtm_dbConnection As New ADODB.Connection
    Dim rtm_rsRecordset As New ADODB.Recordset
    Dim tmpCol As New RefTableCollectionClass
    Dim lCount As Long
    Dim extDataHold As Variant
    Dim extDataHoldCol As Collection
    Dim dbStateAtStart As ADODB.ObjectStateEnum
    
    
    strSQL = GenericSelectAllStatement(TargetRefTable)
    
    Set rtm_dbConnection = RefTableConfig.RefTableDBConnection
    
    dbStateAtStart = rtm_dbConnection.State
    
    If dbStateAtStart <> adStateOpen Then
        rtm_dbConnection.Open (RefTableConfig.RefTableDatabaseLocation)
    Else:
        'TODO: Debug this
        If Not rtm_dbConnection.Properties("Data Source").Value Is Null Then
            If rtm_dbConnection.Properties("Data Source").Value <> RefTableConfig.RefTableDatabaseLocation Then
                MsgBox ("ReferenceTableManagerCore: RefTableDBConnection Data source doesn't match Configuration Database.")
                Exit Function
            End If
        End If
    End If

'    Set rtm_rsRecordset = GetDBRS(rtm_dbConnection)
    
    Set rtm_rsRecordset = New ADODB.Recordset
    
    rtm_rsRecordset.ActiveConnection = rtm_dbConnection

    rtm_rsRecordset.LockType = adLockReadOnly
    
    rtm_rsRecordset.CursorType = adOpenStatic

    rtm_rsRecordset.Source = strSQL

    rtm_rsRecordset.Open
    
'    CombinedConRecordSetPrep rtm_rsRecordset, strSQL
    
    Do Until rtm_rsRecordset.EOF
        Set RefTableData = New RefTableDataClass

        Set extDataHoldCol = New Collection
        For lCount = 0 To rtm_rsRecordset.Fields.Count - 1
            
            Select Case rtm_rsRecordset.Fields(lCount).Name
                Case TargetRefTable.RefID 'Should be a Unique identifier or PK Duplicated values will cause failure when adding to collection.
                    'TODO: Maybe create a way to handle if this field in DB is not Unique.
                    RefTableData.RefTypeID = rtm_rsRecordset.Fields(lCount).Value
                Case TargetRefTable.RefName
                    If Not IsNull(rtm_rsRecordset.Fields(lCount).Value) Then
                        RefTableData.RefTypeName = rtm_rsRecordset.Fields(lCount).Value
                    Else:
                        RefTableData.RefTypeName = ""
                    End If
                Case Else:
                    'Field name is the Key, Field value is the Item
                    If Not IsNull(rtm_rsRecordset.Fields(lCount)) Then
                        RefTableData.RefTypeExtraData.Add rtm_rsRecordset.Fields(lCount).Value, CStr(rtm_rsRecordset.Fields(lCount).Name)
                    Else
                        RefTableData.RefTypeExtraData.Add "", CStr(rtm_rsRecordset.Fields(lCount).Name)
                    End If
            End Select
            
        Next lCount
        
        tmpCol.Add RefTableData
        
        rtm_rsRecordset.MoveNext
    Loop
    
    rtm_rsRecordset.Close
    
    If dbStateAtStart = adStateClosed Then
        rtm_dbConnection.Close
    End If
    
    Set rtm_rsRecordset = Nothing
    Set rtm_dbConnection = Nothing
    
    Set BuildRefTableCollectionClass = tmpCol
    
    
End Function

Private Function GenericSelectAllStatement(ByRef TargetRefTable As IReferenceTableConfig) As String
    If InStr(1, TargetRefTable.RefTableName, " ", vbTextCompare) > 0 Then
        GenericSelectAllStatement = "SELECT * FROM [" & TargetRefTable.RefTableName & "]"
    Else
        GenericSelectAllStatement = "SELECT * FROM " & TargetRefTable.RefTableName
    End If
End Function

Public Function CORE_CheckIfTextExists(ByRef TargetRefTable As IReferenceTableConfig, ByVal TargetText As String) As Boolean
    Dim colToCheck As Collection
    
    Set colToCheck = CORE_GetRefCol(TargetRefTable)
    
    CORE_CheckIfTextExists = RTMInCollection(colToCheck, CStr(TargetText))
            
End Function

Public Function CORE_CheckIfIDExists(ByRef TargetRefTable As IReferenceTableConfig, ByVal TargetID As String) As Boolean
    Dim colToCheck As Collection
    
    Set colToCheck = CORE_GetRefCol(TargetRefTable, rtkt_RefTypeRefID)
    
    CORE_CheckIfIDExists = RTMInCollection(colToCheck, CStr(TargetID))
    
End Function

Public Function CORE_AddNewRefData(ByRef TargetRefTable As IReferenceTableConfig, ByRef NewRefTableData As RefTableDataClass)
    'RefTables with Extra Data (Field count over 2)
    Dim strSQL As String
    Dim rtm_dbConnection As New ADODB.Connection
    Dim rtm_rsRecordset As New ADODB.Recordset
    Dim tmpCol As New Collection
    Dim lCount As Long
    Dim extDataFieldName As Variant
    Dim extDataHold As Variant
    Dim tmpRefTableColData As RefTableCollectionClass
    Dim dbStateAtStart As ADODB.ObjectStateEnum
    
    
    If LastRefTableUpdate(TargetRefTable) > CDate(RefTableBuildDate.Item(TargetRefTable.RefTableName)) Then
        CORE_UpdateRefTableCollections , TargetRefTable
    End If
    'TODO: Is this still a proper check?
    If Not CORE_CheckIfTextExists(TargetRefTable, CStr(NewRefTableData.RefTypeName)) Then
        
        strSQL = GenericSelectAllStatement(TargetRefTable)
    
        Set rtm_dbConnection = RefTableConfig.RefTableDBConnection
        
        dbStateAtStart = rtm_dbConnection.State
    
        If dbStateAtStart <> adStateOpen Then
            rtm_dbConnection.Open (RefTableConfig.RefTableDatabaseLocation)
        Else:
            'TODO: Debug this
            If Not rtm_dbConnection.Properties("Data Source").Value Is Null Then
                If rtm_dbConnection.Properties("Data Source").Value <> RefTableConfig.RefTableDatabaseLocation Then
                    MsgBox ("ReferenceTableManagerCore: RefTableDBConnection Data source doesn't match Configuration Database.")
                    Exit Function
                End If
            End If
        End If

'        Set rtm_rsRecordset = GetDBRS(rtm_dbConnection)

        Set rtm_rsRecordset = New ADODB.Recordset
    
        rtm_rsRecordset.ActiveConnection = rtm_dbConnection

        rtm_rsRecordset.LockType = adLockOptimistic
        
        rtm_rsRecordset.CursorType = adOpenKeyset

        rtm_rsRecordset.Source = strSQL

        rtm_rsRecordset.Open
        
'        CombinedConRecordSetPrep rtm_rsRecordset, strSQL, adOpenKeyset, adLockOptimistic
        
        rtm_rsRecordset.AddNew
        If NewRefTableData.RefTypeID <> -1 Then
            rtm_rsRecordset.Fields(TargetRefTable.RefID).Value = NewRefTableData.RefTypeID
        End If
        rtm_rsRecordset.Fields(TargetRefTable.RefName).Value = NewRefTableData.RefTypeName
        
        For Each extDataFieldName In NewRefTableData.RefTypeExtraData.GetCollectionByExtraDataType(rtedt_FieldNames)
            'TODO: May need to add error handling code for Field changes mid process... though this should be mitigated
            'by using a system that creates an empty ReftableDataClass to put the new data into. (GetConfiguredRefTableDataClass)
            rtm_rsRecordset.Fields(extDataFieldName).Value = NewRefTableData.RefTypeExtraData.Item(extDataFieldName, rtedt_FieldValue)
        Next extDataFieldName
        
        
        rtm_rsRecordset.Update
        
        
        LastRefTableUpdate TargetRefTable, True
        
        rtm_rsRecordset.Close
        
        If dbStateAtStart = adStateClosed Then
            rtm_dbConnection.Close
        End If
        
        Set rtm_rsRecordset = Nothing
        Set rtm_dbConnection = Nothing
        
'        If blnIgnoreRefUpdateChecks = False Then
        CORE_UpdateRefTableCollections True, TargetRefTable
'        Else:
'            Set tmpRefTableColData = colRefTableCollections.Item(GetTableNameForRefTableEnu(TargetRefTable))
'            tmpRefTableColData.Add NewRefTableData
'        End If
        
    End If
    
    
End Function

Public Function CORE_UpdateExistingRefData(ByRef TargetRefTable As IReferenceTableConfig, ByRef NewRefTableData As RefTableDataClass)
    Dim strSQL As String
    Dim rtm_dbConnection As New ADODB.Connection
    Dim rtm_rsRecordset As New ADODB.Recordset
    Dim tmpCol As New Collection
    Dim lCount As Long
    Dim extDataFieldName As Variant
    Dim extDataHold As Variant
    Dim tmpRefTableColData As RefTableCollectionClass
    Dim dbStateAtStart As ADODB.ObjectStateEnum
    
    If LastRefTableUpdate(TargetRefTable) > CDate(RefTableBuildDate.Item(TargetRefTable.RefTableName)) Then
        CORE_UpdateRefTableCollections , TargetRefTable
    End If
    
    'issues if refID is a string
    If InStr(1, TargetRefTable.RefTableName, " ", vbTextCompare) > 0 Then
        strSQL = "SELECT * FROM [" & TargetRefTable.RefTableName & "] Where RefID = " & NewRefTableData.RefTypeID
    Else:
        strSQL = "SELECT * FROM " & TargetRefTable.RefTableName & " Where RefID = " & NewRefTableData.RefTypeID
    End If

    Set rtm_dbConnection = RefTableConfig.RefTableDBConnection
    
    dbStateAtStart = rtm_dbConnection.State

    If dbStateAtStart <> adStateOpen Then
        rtm_dbConnection.Open (RefTableConfig.RefTableDatabaseLocation)
    Else:
        'TODO: Debug this
        If Not rtm_dbConnection.Properties("Data Source").Value Is Null Then
            If rtm_dbConnection.Properties("Data Source").Value <> RefTableConfig.RefTableDatabaseLocation Then
                MsgBox ("ReferenceTableManagerCore: RefTableDBConnection Data source doesn't match Configuration Database.")
                Exit Function
            End If
        End If
    End If

    Set rtm_rsRecordset = New ADODB.Recordset
    
    rtm_rsRecordset.ActiveConnection = rtm_dbConnection

    rtm_rsRecordset.LockType = adLockOptimistic
    
    rtm_rsRecordset.CursorType = adOpenKeyset

    rtm_rsRecordset.Source = strSQL

    rtm_rsRecordset.Open

'    CombinedConRecordSetPrep rtm_rsRecordset, strSQL, adOpenKeyset, adLockOptimistic
    
    If rtm_rsRecordset.RecordCount = 0 Then
        CORE_AddNewRefData TargetRefTable, NewRefTableData
    ElseIf rtm_rsRecordset.RecordCount = 1 Then
    
        If rtm_rsRecordset.Fields(TargetRefTable.RefID).Value <> NewRefTableData.RefTypeID Then
            rtm_rsRecordset.Fields(TargetRefTable.RefID).Value = NewRefTableData.RefTypeID
        End If
        
        
        If rtm_rsRecordset.Fields(TargetRefTable.RefName).Value <> NewRefTableData.RefTypeName Then
            rtm_rsRecordset.Fields(TargetRefTable.RefName).Value = NewRefTableData.RefTypeName
        End If
        
        For Each extDataFieldName In NewRefTableData.RefTypeExtraData.GetCollectionByExtraDataType(rtedt_FieldNames)
            'TODO: May need to add error handling code for Field changes mid process... though this should be mitigated
            'by using a system that creates an empty ReftableDataClass to put the new data into. (GetConfiguredRefTableDataClass)
            If rtm_rsRecordset.Fields(extDataFieldName).Value <> NewRefTableData.RefTypeExtraData.Item(extDataFieldName, rtedt_FieldValue) Then
                rtm_rsRecordset.Fields(extDataFieldName).Value = NewRefTableData.RefTypeExtraData.Item(extDataFieldName, rtedt_FieldValue)
            End If
        Next extDataFieldName
        
        rtm_rsRecordset.Update
        
        LastRefTableUpdate TargetRefTable, True
    ElseIf rtm_rsRecordset.RecordCount >= 2 Then
        'Shouldn't happen but could depending on config or improper use
    End If
    
    rtm_rsRecordset.Close

    If dbStateAtStart = adStateClosed Then
        rtm_dbConnection.Close
    End If
    
    Set rtm_rsRecordset = Nothing
    Set rtm_dbConnection = Nothing
    
'   If blnIgnoreRefUpdateChecks = False Then
        CORE_UpdateRefTableCollections True, TargetRefTable
'   Else:
'       Set tmpRefTableColData = colRefTableCollections.Item(GetTableNameForRefTableEnu(TargetRefTable))
'       tmpRefTableColData.Add NewRefTableData
'   End If
    
End Function

Private Function LastRefTableUpdate(ByRef TargetRefTable As IReferenceTableConfig, Optional ByVal UpdateDate As Boolean = False) As Date
    'Returns the Date in the ApplicationDataStore pertaining to the TargetRefTable supplied, or if UpdateDate = True
    'updates the date to the current date and time and returns that. Also returns current date if ApplicationDataStore's
    'ElementDataStore field is Null.
    '
    'Handles creation of new entries in the ApplicationDataStore table for initial load of reftables, or as new ones
    'get created in the DB and added to the code at a later date. (This does not create new RefTables. Only adds there
    'ApplicationElement and ElementDataStore data to the fields in ApplicationDataStore.)
    
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim rtnDateTime As Date
    Dim strSQL As String
    Dim dbStateAtStart As ADODB.ObjectStateEnum
    Dim dbCreateAppElmTable As ADODB.Command
    
UpdateStart:

    strSQL = "Select * from ApplicationDataStore where ApplicationElement = '" & TargetRefTable.RefTableName & "Update'"
    
    Set DBCon = RefTableConfig.RefTableDBConnection

    dbStateAtStart = DBCon.State

    If dbStateAtStart <> adStateOpen Then
        DBCon.Open (RefTableConfig.RefTableDatabaseLocation)
    Else:
        'TODO: Debug this
        If Not DBCon.Properties("Data Source").Value = vbNullString Then
            If DBCon.Properties("Data Source").Value <> RefTableConfig.RefTableDatabaseLocation Then
                MsgBox ("ReferenceTableManagerCore: RefTableDBConnection Data source doesn't match Configuration Database.")
                Exit Function
            End If
        End If
    End If

    Set dbRS = New ADODB.Recordset
    
    dbRS.ActiveConnection = DBCon

    dbRS.Source = strSQL
    
    dbRS.CursorType = adOpenKeyset
    
    dbRS.LockType = adLockOptimistic
    
    On Error GoTo TableMissing
    dbRS.Open
    On Error GoTo 0
    
'    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockOptimistic
    
    'Adds new Reftable entry to ApplicationDataStore if not found
    If dbRS.RecordCount = 0 Then
        dbRS.AddNew
        dbRS.Fields("ApplicationElement").Value = TargetRefTable.RefTableName & "Update"
        dbRS.Update
    End If
    
    'Checks for date and/or updates date
    If UpdateDate = False Then
        If Not IsNull(dbRS.Fields("ElementDataStore").Value) Then
            rtnDateTime = dbRS.Fields("ElementDataStore").Value
        Else:
            rtnDateTime = RTM_EMPTY_DATE
        End If
    Else:
        dbRS.Fields("ElementDataStore").Value = DateTime.Now
        dbRS.Update
        rtnDateTime = dbRS.Fields("ElementDataStore").Value
    End If
        
    dbRS.Close
    
    If dbStateAtStart = adStateClosed Then
        DBCon.Close
    End If
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    LastRefTableUpdate = rtnDateTime
    
    Exit Function
    
TableMissing:
    'Check error
    'Handle if Table error

    If Err.Number <> -2147217865 Then
        MsgBox ("Unknown ADODB Error: " & Err.Number & "; " & Err.Description)
        On Error GoTo 0
        Exit Function
    End If

    'Generate Table
    On Error GoTo 0
    
    If dbRS.State = adStateOpen Then
        dbRS.Close
    End If
    
    Set dbCreateAppElmTable = New ADODB.Command
    
    dbCreateAppElmTable.ActiveConnection = DBCon
    dbCreateAppElmTable.CommandText = "CREATE TABLE ApplicationDataStore(ID AUTOINCREMENT PRIMARY KEY, ApplicationElement VARCHAR, ElementDataStore VARCHAR)"
    dbCreateAppElmTable.CommandType = adCmdText
    
    dbCreateAppElmTable.Execute
    
'    dbRS.Source = "CREATE TABLE ApplicationDataStore(ID AUTOINCREMENT PRIMARY KEY, ApplicationElement VARCHAR, ElementDataStore VARCHAR)"
    
'    dbRS.Open
'
'    dbRS.Close

    Resume UpdateStart

End Function

Public Function CORE_GetConfiguredRefTableDataClass(ByRef TargetRefTable As IReferenceTableConfig) As RefTableDataClass
    'Returns an empty RefTableDataClass that has been configured with the fields from the TargetRefTable
    'for use with adding new RefTableData
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim lCount As Long
    Dim rtnRefTabDataCls As RefTableDataClass
    Dim strSQL As String
    Dim dbStateAtStart As ADODB.ObjectStateEnum
    
    strSQL = GenericSelectAllStatement(TargetRefTable)
    
    Set DBCon = RefTableConfig.RefTableDBConnection

    dbStateAtStart = DBCon.State

    If dbStateAtStart <> adStateOpen Then
        DBCon.Open (RefTableConfig.RefTableDatabaseLocation)
    Else:
        'TODO: Debug this
        If Not DBCon.Properties("Data Source").Value Is Null Then
            If DBCon.Properties("Data Source").Value <> RefTableConfig.RefTableDatabaseLocation Then
                MsgBox ("ReferenceTableManagerCore: RefTableDBConnection Data source doesn't match Configuration Database.")
                Exit Function
            End If
        End If
    End If
    
    Set dbRS = New ADODB.Recordset
    
    dbRS.ActiveConnection = DBCon

    dbRS.Source = strSQL
    
    dbRS.CursorType = adOpenKeyset
    
    dbRS.LockType = adLockOptimistic

    dbRS.Open

'    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockOptimistic
    
    Set rtnRefTabDataCls = New RefTableDataClass
    
    For lCount = 0 To dbRS.Fields.Count - 1
        If dbRS.Fields(lCount).Name <> TargetRefTable.RefID And dbRS.Fields(lCount).Name <> TargetRefTable.RefName Then
            rtnRefTabDataCls.RefTypeExtraData.Add "", CStr(dbRS.Fields(lCount).Name)
        End If
    Next lCount
    
    dbRS.Close

    If dbStateAtStart = adStateClosed Then
        DBCon.Close
    End If
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    Set CORE_GetConfiguredRefTableDataClass = rtnRefTabDataCls
    
End Function

Public Function CORE_IgnoreRefUpdateChecks(ByVal Ignore As Boolean)
    blnIgnoreRefUpdateChecks = Ignore
End Function

Public Function CORE_IsIgnoringUpdates() As Boolean
    CORE_IsIgnoringUpdates = blnIgnoreRefUpdateChecks
End Function

Private Function RTMInCollection(colCollection As Collection, ByVal strItemToCheck As String) As Boolean
    'Compares value provided to collection provided key's refrence. On error (no item found) returns false, else true
    ' a.k.a Key found.
    On Error GoTo HandleError
    'On Error Resume Next
    Dim var As Variant
    var = colCollection(strItemToCheck)
    RTMInCollection = True
    Exit Function
HandleError:
    If Err.Number <> 438 Then
        RTMInCollection = False
    Else:
        RTMInCollection = True
    End If
End Function
