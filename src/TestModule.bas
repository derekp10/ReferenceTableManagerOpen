Attribute VB_Name = "TestModule"
Option Explicit

Public Sub BasicTest()
    Dim tmpRTM As TestReferenceTableManager
    Dim rtnVal As String
    Dim tmpRTDC As RefTableDataClass
    Dim tmpRTED As RefTableExtraDataCollection
    
    Set tmpRTM = New TestReferenceTableManager
    
    tmpRTM.TestIgnoreRefUpdateChecks False
    
    Set tmpRTDC = tmpRTM.TestGetConfiguredRefTableDataClass(rte_TestRefTable)
    
'    tmpRTDC.RefTypeName = "Test1"
'    tmpRTDC.RefTypeExtraData.SetValueForFieldName "RefExtra", "RefTest1"
'
'    tmpRTM.AddNewRefData rte_TestRefTable, tmpRTDC
    
    'Test existing data load
    
    Debug.Print (tmpRTM.TestGetStringFromID(rte_TestRefTable, CStr("1")))
    
    Debug.Print (tmpRTM.TestGetIDFromString(rte_TestRefTable, "Test1"))
    
    Set tmpRTED = tmpRTM.TestGetExtraDataFromID(rte_TestRefTable, CStr("1"))
    
    Debug.Print (tmpRTED.GetValueForFieldName("RefExtra"))
    
    'Test Add New Functions
    
'    tmpRTDC.RefTypeName = "Test2"
'    tmpRTDC.RefTypeExtraData.SetValueForFieldName "RefExtra", "RefExtraTest2"
'
'    tmpRTM.TestAddNewRefData rte_TestRefTable, tmpRTDC
'
'    Debug.Print (tmpRTM.TestGetStringFromID(rte_TestRefTable, CStr("2")))
'
'    Debug.Print (tmpRTM.TestGetIDFromString(rte_TestRefTable, "Test2"))
'
'    Set tmpRTED = tmpRTM.TestGetExtraDataFromID(rte_TestRefTable, CStr("2"))
'
'    Debug.Print (tmpRTED.GetValueForFieldName("RefExtra"))
'
'
'    Set tmpRTDC = tmpRTM.TestGetConfiguredRefTableDataClass(rte_TestRefTableQuery)
    
    
    'Test existing data load
    
'    Debug.Print (tmpRTM.TestGetStringFromID(rte_TestRefTable, CStr("1")))
'
'    Debug.Print (tmpRTM.TestGetIDFromString(rte_TestRefTable, "Test1, RefTest1"))
'
'    Set tmpRTED = tmpRTM.TestGetExtraDataFromID(rte_TestRefTable, CStr("1"))
'
''    Debug.Print (tmpRTED.GetValueForFieldName("RefExtra"))
'
    Debug.Print (tmpRTM.TestIsIgnoringUpdates)
    
    
    
    Debug.Print (rtnVal)
End Sub

Public Sub BasicAddTest()
    'Delete record entry from the TestRefTable and compact/repair database after running this test if you want to run it again.
    Dim tmpRTM As TestReferenceTableManager
    Dim rtnVal As String
    Dim tmpRTDC As RefTableDataClass
    Dim tmpRTED As RefTableExtraDataCollection

    Set tmpRTM = New TestReferenceTableManager
    
    Set tmpRTDC = tmpRTM.TestGetConfiguredRefTableDataClass(rte_TestRefTable)

    tmpRTDC.RefTypeName = "Test2"
    tmpRTDC.RefTypeExtraData.SetValueForFieldName "RefExtra", "RefExtraTest2"

    tmpRTM.TestAddNewRefData rte_TestRefTable, tmpRTDC
    
    Debug.Print (tmpRTM.TestGetStringFromID(rte_TestRefTable, CStr("2")))
    
    Debug.Print (tmpRTM.TestGetIDFromString(rte_TestRefTable, "Test2"))
    
    Set tmpRTED = tmpRTM.TestGetExtraDataFromID(rte_TestRefTable, CStr("2"))
    
    Debug.Print (tmpRTED.GetValueForFieldName("RefExtra"))
    

End Sub

Public Sub BasicQueryTest()
    Dim tmpRTM As TestReferenceTableManager
    Dim rtnVal As String
    Dim tmpRTDC As RefTableDataClass
    Dim tmpRTED As RefTableExtraDataCollection
    
    Set tmpRTM = New TestReferenceTableManager
    
    Set tmpRTDC = tmpRTM.TestGetConfiguredRefTableDataClass(rte_TestRefTableQuery)
    
    'Test existing data load
    
    Debug.Print (tmpRTM.TestGetStringFromID(rte_TestRefTableQuery, CStr("1")))
    
    Debug.Print (tmpRTM.TestGetIDFromString(rte_TestRefTableQuery, "Test1, RefTest1"))
    
    Set tmpRTED = tmpRTM.TestGetExtraDataFromID(rte_TestRefTableQuery, CStr("1"))
    
'    Debug.Print (tmpRTED.GetValueForFieldName("RefExtra"))
End Sub

Public Sub BasicQueryAddTest()
    'For testing only, this will fail as Query based adds are not currently supported
    Dim tmpRTM As TestReferenceTableManager
    Dim rtnVal As String
    Dim tmpRTDC As RefTableDataClass
    Dim tmpRTED As RefTableExtraDataCollection

    Set tmpRTM = New TestReferenceTableManager
    
    Set tmpRTDC = tmpRTM.TestGetConfiguredRefTableDataClass(rte_TestRefTableQuery)

    tmpRTDC.RefTypeName = "Test2"
    tmpRTDC.RefTypeExtraData.SetValueForFieldName "RefExtra", "RefExtraTest2"
    
    tmpRTM.TestAddNewRefData rte_TestRefTableQuery, tmpRTDC
    
    Debug.Print (tmpRTM.TestGetStringFromID(rte_TestRefTableQuery, CStr("2")))
    
    Debug.Print (tmpRTM.TestGetIDFromString(rte_TestRefTableQuery, "Test2, RefExtraTest2"))
    
    Set tmpRTED = tmpRTM.TestGetExtraDataFromID(rte_TestRefTableQuery, CStr("2"))
    
    'Debug.Print (tmpRTED.GetValueForFieldName("RefExtra"))
End Sub
