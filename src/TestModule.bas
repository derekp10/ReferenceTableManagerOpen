Attribute VB_Name = "TestModule"
Option Explicit

Public Sub BasicTest()
    Dim tmpRTM As ReferenceTableManager
    Dim rtnVal As String
    Dim tmpRTDC As RefTableDataClass
    Dim tmpRTED As RefTableExtraDataCollection
    
    Set tmpRTM = New ReferenceTableManager
    
    tmpRTM.IgnoreRefUpdateChecks False
    
    Set tmpRTDC = tmpRTM.GetConfiguredRefTableDataClass(rte_TestRefTable)
    
'    tmpRTDC.RefTypeName = "Test1"
'    tmpRTDC.RefTypeExtraData.SetValueForFieldName "RefExtra", "RefTest1"
'
'    tmpRTM.AddNewRefData rte_TestRefTable, tmpRTDC
    
    Debug.Print (tmpRTM.GetStringFromID(rte_TestRefTable, CStr("1")))
    
    Debug.Print (tmpRTM.GetIDFromString(rte_TestRefTable, "Test1"))
    
    Set tmpRTED = tmpRTM.GetExtraDataFromID(rte_TestRefTable, CStr("1"))
    
    Debug.Print (tmpRTED.GetValueForFieldName("RefExtra"))
    
    
    
    Debug.Print (tmpRTM.IsIgnoringUpdates)
    
    
    
    Debug.Print (rtnVal)
End Sub
