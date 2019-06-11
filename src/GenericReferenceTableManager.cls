VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericReferenceTableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'GenericReferenceTableManager
'By Derek Pierce (12/6/2016)
'Ver. 1.0 12/6/2016
'Ver. 1.3 6/10/2019

Implements IReferenceTableManagerConfig

'Used to configure and manage the look up data for different RefTables in the Database
'!IMPORTANT!: All sections with the "!IMPORTANT!: This section must be set up prior to use with a new project."
'line must be configured inorder for this system to work, as well as the variables below.
'It would also be a good idea to use Find and Replace to Change all "Generic_" references to those that
'identify what this implementation will work for. You should also change the name of this class to reflect what Generic_
'was changed to. A good suggestion would be the name of your project. That way it is easier
'to add the reference data from another database.
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
Public Enum Generic_RefTableEnu
    [_First] = 1
        rte_GENERIC_PLACEHOLDERDONOTUSE = 0 'DO NOT USE THIS. CODE WILL BREAK. INTERNAL CLASS USE ONLY
        'rte_DebugRefTable = 1 'Example Entry
    [_Last] = 1 'Update to number of last entry in Enumerator
End Enum


Private RefTableMngCORE As New ReferenceTableManagerCore
Private colRefTableCollections As Collection

Public Function Generic_GetTableConfigForRefTableEnu(ByVal TargetRefTable As Generic_RefTableEnu) As IReferenceTableConfig
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
    Set Generic_GetTableConfigForRefTableEnu = objRefTableConfg
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

'Generic Code (Don't forget to replace Generic_ with an actual name if using more than one config)

Private Sub Class_Initialize()
    Set RefTableMngCORE = New ReferenceTableManagerCore
    RefTableMngCORE.SetRefTableConfig = Me
    Set colRefTableCollections = New Collection
    Generic_UpdateRefTableCollections
End Sub

Public Function Generic_GetRefCol(ByVal TargetRefTable As Generic_RefTableEnu, Optional KeyType As RefTableKeyType = rtkt_RefTypeRefName) As Collection
    Set Generic_GetRefCol = RefTableMngCORE.CORE_GetRefCol(Generic_GetTableConfigForRefTableEnu(TargetRefTable), KeyType)
End Function

Public Function Generic_UpdateRefTableCollections(Optional ByVal ForceUpdate As Boolean = False, Optional ByVal TargetRefTable As Generic_RefTableEnu = rte_GENERIC_PLACEHOLDERDONOTUSE)
    Dim RefEnuNum As Long

    If TargetRefTable = rte_GENERIC_PLACEHOLDERDONOTUSE Then
        For RefEnuNum = Generic_RefTableEnu.[_First] To Generic_RefTableEnu.[_Last]
            RefTableMngCORE.CORE_UpdateRefTableCollections True, Generic_GetTableConfigForRefTableEnu(RefEnuNum)
        Next RefEnuNum
    Else:
        RefTableMngCORE.CORE_UpdateRefTableCollections ForceUpdate, Generic_GetTableConfigForRefTableEnu(TargetRefTable)
    End If

End Function

Public Function Generic_GetIDFromString(ByVal TargetRefTable As Generic_RefTableEnu, strTarget As String) As Long
    ''Using a string value and a reference table to look up, returns the ID for that string
    ''Key lookups are CASE SENSITIVE!!
    Generic_GetIDFromString = RefTableMngCORE.CORE_GetIDFromString(Generic_GetTableConfigForRefTableEnu(TargetRefTable), strTarget)
End Function

Public Function Generic_GetStringFromID(ByVal TargetRefTable As Generic_RefTableEnu, lngTarget As Variant) As String
    ''Using a long(variant) value and a reference table to look up, returns the String value for that ID
    Generic_GetStringFromID = RefTableMngCORE.CORE_GetStringFromID(Generic_GetTableConfigForRefTableEnu(TargetRefTable), lngTarget)
End Function

Public Function Generic_GetExtraDataFromID(ByVal TargetRefTable As Generic_RefTableEnu, lngTarget As Variant) As RefTableExtraDataCollection
    ''Gets any extra data that may be present for the Reftable. (Forward comments etc..)
    Set Generic_GetExtraDataFromID = RefTableMngCORE.CORE_GetExtraDataFromID(Generic_GetTableConfigForRefTableEnu(TargetRefTable), lngTarget)
End Function

Public Function Generic_CheckIfTextExists(ByVal TargetRefTable As Generic_RefTableEnu, ByVal TargetText As String) As Boolean
    Generic_CheckIfTextExists = RefTableMngCORE.CORE_CheckIfTextExists(Generic_GetTableConfigForRefTableEnu(TargetRefTable), TargetText)
End Function

Public Function Generic_CheckIfIDExists(ByVal TargetRefTable As Generic_RefTableEnu, ByVal TargetID As String) As Boolean
    Generic_CheckIfIDExists = RefTableMngCORE.CORE_CheckIfIDExists(Generic_GetTableConfigForRefTableEnu(TargetRefTable), TargetID)
End Function

Public Function Generic_AddNewRefData(ByVal TargetRefTable As Generic_RefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    RefTableMngCORE.CORE_AddNewRefData Generic_GetTableConfigForRefTableEnu(TargetRefTable), NewRefTableData
End Function

Public Function Generic_UpdateExistingRefData(ByVal TargetRefTable As Generic_RefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    RefTableMngCORE.CORE_UpdateExistingRefData Generic_GetTableConfigForRefTableEnu(TargetRefTable), NewRefTableData
End Function

Public Function Generic_GetConfiguredRefTableDataClass(ByVal TargetRefTable As Generic_RefTableEnu) As RefTableDataClass
    Set Generic_GetConfiguredRefTableDataClass = RefTableMngCORE.CORE_GetConfiguredRefTableDataClass(Generic_GetTableConfigForRefTableEnu(TargetRefTable))
End Function

Public Function Generic_IgnoreRefUpdateChecks(ByVal Ignore As Boolean)
    RefTableMngCORE.CORE_IgnoreRefUpdateChecks Ignore
End Function

Public Function Generic_IsIgnoringUpdates() As Boolean
    Generic_IsIgnoringUpdates = RefTableMngCORE.CORE_IsIgnoringUpdates
End Function
