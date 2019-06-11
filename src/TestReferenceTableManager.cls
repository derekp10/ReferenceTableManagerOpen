VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestReferenceTableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TestReferenceTableManager
'By Derek Pierce (12/6/2016)
'Ver. 1.0 12/6/2016
'Ver. 1.3 6/10/2019

Implements IReferenceTableManagerConfig

'Used to configure and manage the look up the data for different RefTables in the Database
'!IMPORTANT!: All sections with the "!IMPORTANT!: This section must be set up prior to use with a new project."
'line must be configured inorder for this system to work, as well as the variables below.
'It would also be a good idea to use Find and Replace to Change all "Generic_" references to those that
'identify what this implementation will work for. you should also change the name of this class to reflect what Generic_
'was changed to. A good suggestion would be the name of your project. This way it is easier
'to add a the reference data from another database.
'
'The rest of the code should work without modification, as long as all required classes are included.
'Those are:
'GenericReferenceTableManager (This Class)
'ReferenceTableManagerCore
'RefTableCollectionClass
'RefTableDataClass (Requires ISearchable interface from the SearchableComboBox Code folder)
'RefTableExtraDataCollection
'ISearchable (Unless the Interface Implement call is disabled in RefTableDataClass)

'Enumerator used to hard define the tables used/managed by this code.
'Makes referencing tables easier in the "GetXFromY" function lookups
'If enum has no entries, then references will break and not work properly.
'!IMPORTANT!: This section must be set up prior to use with a new project.
'Replace Generic with the name of the project/database this is being used for.
'Use Find and Replace to update all references in code.
Public Enum TestRefTableEnu
    [_First] = 1
        rte_TestPLACEHOLDERDONOTUSE = 0 'DO NOT USE THIS. CODE WILL BREAK. INTERNAL CLASS USE ONLY
        'rte_DebugRefTable = 1 'Example Entry
        rte_TestRefTable = 1
        rte_TestRefTableQuery
    [_Last] = 2 'Update to number of last entry in Enumerator
End Enum


Private RefTableMngCORE As New ReferenceTableManagerCore
Private colRefTableCollections As Collection

Public Function TestGetTableConfigForRefTableEnu(ByVal TargetRefTable As TestRefTableEnu) As IReferenceTableConfig
    '!IMPORTANT!: This section must be set up prior to use with a new project.
    'NOTICE: Add RefTable Names here.
    'If table(s) are not added to code it returns null string or error
    Dim objRefTableConfg As GenericReferenceTableConfig

    Set objRefTableConfg = New GenericReferenceTableConfig

    'Example Entry
'    If TargetRefTable = rte_DebugRefTable Then 'Specify the Enumerator
'        'Specify the data returned for the enumerator
'        objRefTableConfg.SetConfigValues "DebugRefTable", "DebugRefID", "DebugRefName"
'        GoTo TableFound:
'    End If

    If TargetRefTable = rte_TestRefTable Then 'Specify the Enumerator
        'Specify the data returned for the enumerator
'        objRefTableConfg.SetConfigValues "TestRefTable", "TestRefID", "TestRefName"
        objRefTableConfg.SetConfigValues "TestRefTable"
        GoTo TableFound:
    End If
    
    If TargetRefTable = rte_TestRefTableQuery Then 'Specify the Enumerator
        'Specify the data returned for the enumerator
'        objRefTableConfg.SetConfigValues "TestRefTable", "TestRefID", "TestRefName"
        objRefTableConfg.SetConfigValues "TestRefTableQuery", , "CombinedRefName"
        GoTo TableFound:
    End If

TableFound:
    Set TestGetTableConfigForRefTableEnu = objRefTableConfg
End Function

'Interface Implementations

Private Property Let IReferenceTableManagerConfig_RefTableCollections(RHS As Collection)
    Set colRefTableCollections = RHS
End Property

Private Property Get IReferenceTableManagerConfig_RefTableCollections() As Collection
    Set IReferenceTableManagerConfig_RefTableCollections = colRefTableCollections
End Property

'Database Location must be supplied.
Private Property Get IReferenceTableManagerConfig_RefTableDatabaseLocation() As String
    '!IMPORTANT!: This section must be set up prior to use with a new project.
    IReferenceTableManagerConfig_RefTableDatabaseLocation = DB_LOC & DB_NAME 'Database Location
End Property

Private Property Get IReferenceTableManagerConfig_RefTableDBConnection() As ADODB.Connection
    '!IMPORTANT!: This section must be set up prior to use with a new project.
    'Supply a configured Active X Data Object Connection Object.
    Set IReferenceTableManagerConfig_RefTableDBConnection = GetDBCon
End Property

'Generic Code (Don't forget to replace Generic_ with an actual name if using more than one config)

Private Sub Class_Initialize()
    Set RefTableMngCORE = New ReferenceTableManagerCore
    RefTableMngCORE.SetRefTableConfig = Me
    Set colRefTableCollections = New Collection
    TestUpdateRefTableCollections
End Sub

Public Function TestGetRefCol(ByVal TargetRefTable As TestRefTableEnu, Optional KeyType As RefTableKeyType = rtkt_RefTypeRefName) As Collection
    Set TestGetRefCol = RefTableMngCORE.CORE_GetRefCol(TestGetTableConfigForRefTableEnu(TargetRefTable), KeyType)
End Function

Public Function TestUpdateRefTableCollections(Optional ByVal ForceUpdate As Boolean = False, Optional ByVal TargetRefTable As TestRefTableEnu = rte_TestPLACEHOLDERDONOTUSE)
    Dim RefEnuNum As Long

    If TargetRefTable = rte_TestPLACEHOLDERDONOTUSE Then
        For RefEnuNum = TestRefTableEnu.[_First] To TestRefTableEnu.[_Last]
            RefTableMngCORE.CORE_UpdateRefTableCollections True, TestGetTableConfigForRefTableEnu(RefEnuNum)
        Next RefEnuNum
    Else:
        RefTableMngCORE.CORE_UpdateRefTableCollections ForceUpdate, TestGetTableConfigForRefTableEnu(TargetRefTable)
    End If

End Function

Public Function TestGetIDFromString(ByVal TargetRefTable As TestRefTableEnu, strTarget As String) As Long
    ''Using a string value and a reference table to look up, returns the ID for that string
    ''Key lookups are CASE SENSITIVE!!
    TestGetIDFromString = RefTableMngCORE.CORE_GetIDFromString(TestGetTableConfigForRefTableEnu(TargetRefTable), strTarget)
End Function

Public Function TestGetStringFromID(ByVal TargetRefTable As TestRefTableEnu, lngTarget As Variant) As String
    ''Using a long(variant) value and a reference table to look up, returns the String value for that ID
    TestGetStringFromID = RefTableMngCORE.CORE_GetStringFromID(TestGetTableConfigForRefTableEnu(TargetRefTable), lngTarget)
End Function

Public Function TestGetExtraDataFromID(ByVal TargetRefTable As TestRefTableEnu, lngTarget As Variant) As RefTableExtraDataCollection
    ''Gets any extra data that may be present for the Reftable. (Forward comments etc..)
    Set TestGetExtraDataFromID = RefTableMngCORE.CORE_GetExtraDataFromID(TestGetTableConfigForRefTableEnu(TargetRefTable), lngTarget)
End Function

Public Function TestCheckIfTextExists(ByVal TargetRefTable As TestRefTableEnu, ByVal TargetText As String) As Boolean
    TestCheckIfTextExists = RefTableMngCORE.CORE_CheckIfTextExists(TestGetTableConfigForRefTableEnu(TargetRefTable), TargetText)
End Function

Public Function TestCheckIfIDExists(ByVal TargetRefTable As TestRefTableEnu, ByVal TargetID As String) As Boolean
    TestCheckIfIDExists = RefTableMngCORE.CORE_CheckIfIDExists(TestGetTableConfigForRefTableEnu(TargetRefTable), TargetID)
End Function

Public Function TestAddNewRefData(ByVal TargetRefTable As TestRefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    RefTableMngCORE.CORE_AddNewRefData TestGetTableConfigForRefTableEnu(TargetRefTable), NewRefTableData
End Function

Public Function TestUpdateExistingRefData(ByVal TargetRefTable As TestRefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    RefTableMngCORE.CORE_UpdateExistingRefData TestGetTableConfigForRefTableEnu(TargetRefTable), NewRefTableData
End Function

Public Function TestGetConfiguredRefTableDataClass(ByVal TargetRefTable As TestRefTableEnu) As RefTableDataClass
    Set TestGetConfiguredRefTableDataClass = RefTableMngCORE.CORE_GetConfiguredRefTableDataClass(TestGetTableConfigForRefTableEnu(TargetRefTable))
End Function

Public Function TestIgnoreRefUpdateChecks(ByVal Ignore As Boolean)
    RefTableMngCORE.CORE_IgnoreRefUpdateChecks Ignore
End Function

Public Function TestIsIgnoringUpdates() As Boolean
    TestIsIgnoringUpdates = RefTableMngCORE.CORE_IsIgnoringUpdates
End Function

