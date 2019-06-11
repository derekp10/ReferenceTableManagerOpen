VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceTableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ReferenceTableManager
'By Derek Pierce (12/6/2016)
'Ver. 1.0 12/6/2016 Based off of GenericReferenceTableManager
'Ver. 1.3 6/10/2019

Implements IReferenceTableManagerConfig

'THIS IS AN IMPLEMENTATION OF THE ORIGINAL CODE IN ORDER FOR IT TO BE USED IN PLACE ON
'PROJECTS THAT WERE DONE BEFORE THE 1.2 REWORK. IF ADDING A NEW EXTERNAL REFERENCE TABLE
'USE GenericReferenceTableManager AND FOLLOW IT'S INSTRUCTIONS.

'Used to configure and manage the look up the data for different RefTables in the Database
'!IMPORTANT!: All sections with the "!IMPORTANT!: This section must be set up prior to use with a new project."
'line must be configured inorder for this system to work, as well as the variables below.
'
'The rest of the code should work without modification, as long as all required classes are included.
'Those are:
'ReferenceTableManager (This Class)
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
Public Enum RefTableEnu
    [_First] = 1
        rte_PLACEHOLDERDONOTUSE = 0 'DO NOT USE THIS. CODE WILL BREAK. INTERNAL CLASS USE ONLY
        'rte_DebugRefTable = 1 'Example Entry
    [_Last] = 1 'Update to number of last entry in Enumerator
End Enum


Private RefTableMngCORE As New ReferenceTableManagerCore
Private colRefTableCollections As Collection

Public Function GetTableConfigForRefTableEnu(ByVal TargetRefTable As RefTableEnu) As IReferenceTableConfig
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

TableFound:
    Set GetTableConfigForRefTableEnu = objRefTableConfg
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
    IReferenceTableManagerConfig_RefTableDatabaseLocation = "" 'Database Location
End Property

Private Property Get IReferenceTableManagerConfig_RefTableDBConnection() As ADODB.Connection
    '!IMPORTANT!: This section must be set up prior to use with a new project.
    'Supply a configured Active X Data Object Connection Object.
    Set IReferenceTableManagerConfig_RefTableDBConnection = ""
End Property

'Generic Code (Wrapper for old function calls in 1.1 version of ReferenceTableManager/ReferenceTableManagerCore)

Private Sub Class_Initialize()
    Set RefTableMngCORE = New ReferenceTableManagerCore
    RefTableMngCORE.SetRefTableConfig = Me
    Set colRefTableCollections = New Collection
    UpdateRefTableCollections
End Sub

Public Function GetRefCol(ByVal TargetRefTable As RefTableEnu, Optional KeyType As RefTableKeyType = rtkt_RefTypeRefName) As Collection
    Set GetRefCol = RefTableMngCORE.CORE_GetRefCol(GetTableConfigForRefTableEnu(TargetRefTable), KeyType)
End Function

Public Function UpdateRefTableCollections(Optional ByVal ForceUpdate As Boolean = False, Optional ByVal TargetRefTable As RefTableEnu = rte_PLACEHOLDERDONOTUSE)
    Dim RefEnuNum As Long

    If TargetRefTable = rte_PLACEHOLDERDONOTUSE Then
        For RefEnuNum = RefTableEnu.[_First] To RefTableEnu.[_Last]
            RefTableMngCORE.CORE_UpdateRefTableCollections True, GetTableConfigForRefTableEnu(RefEnuNum)
        Next RefEnuNum
    Else:
        RefTableMngCORE.CORE_UpdateRefTableCollections ForceUpdate, GetTableConfigForRefTableEnu(TargetRefTable)
    End If

End Function

Public Function GetIDFromString(ByVal TargetRefTable As RefTableEnu, strTarget As String) As Long
    ''Using a string value and a reference table to look up, returns the ID for that string
    ''Key lookups are CASE SENSITIVE!!
    GetIDFromString = RefTableMngCORE.CORE_GetIDFromString(GetTableConfigForRefTableEnu(TargetRefTable), strTarget)
End Function

Public Function GetStringFromID(ByVal TargetRefTable As RefTableEnu, lngTarget As Variant) As String
    ''Using a long(variant) value and a reference table to look up, returns the String value for that ID
    GetStringFromID = RefTableMngCORE.CORE_GetStringFromID(GetTableConfigForRefTableEnu(TargetRefTable), lngTarget)
End Function

Public Function GetExtraDataFromID(ByVal TargetRefTable As RefTableEnu, lngTarget As Variant) As RefTableExtraDataCollection
    ''Gets any extra data that may be present for the Reftable. (Forward comments etc..)
    Set GetExtraDataFromID = RefTableMngCORE.CORE_GetExtraDataFromID(GetTableConfigForRefTableEnu(TargetRefTable), lngTarget)
End Function

Public Function CheckIfTextExists(ByVal TargetRefTable As RefTableEnu, ByVal TargetText As String) As Boolean
    CheckIfTextExists = RefTableMngCORE.CORE_CheckIfTextExists(GetTableConfigForRefTableEnu(TargetRefTable), TargetText)
End Function

Public Function CheckIfIDExists(ByVal TargetRefTable As RefTableEnu, ByVal TargetID As String) As Boolean
    CheckIfIDExists = RefTableMngCORE.CORE_CheckIfIDExists(GetTableConfigForRefTableEnu(TargetRefTable), TargetID)
End Function

Public Function AddNewRefData(ByVal TargetRefTable As RefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    RefTableMngCORE.CORE_AddNewRefData GetTableConfigForRefTableEnu(TargetRefTable), NewRefTableData
End Function

Public Function UpdateExistingRefData(ByVal TargetRefTable As RefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    RefTableMngCORE.CORE_UpdateExistingRefData GetTableConfigForRefTableEnu(TargetRefTable), NewRefTableData
End Function

Public Function GetConfiguredRefTableDataClass(ByVal TargetRefTable As RefTableEnu) As RefTableDataClass
    Set GetConfiguredRefTableDataClass = RefTableMngCORE.CORE_GetConfiguredRefTableDataClass(GetTableConfigForRefTableEnu(TargetRefTable))
End Function

Public Function IgnoreRefUpdateChecks(ByVal Ignore As Boolean)
    RefTableMngCORE.CORE_IgnoreRefUpdateChecks Ignore
End Function

Public Function IsIgnoringUpdates() As Boolean
    IsIgnoringUpdates = RefTableMngCORE.CORE_IsIgnoringUpdates
End Function

