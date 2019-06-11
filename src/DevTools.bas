Attribute VB_Name = "DevTools"
Option Explicit

'https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
'With modifications to work with Git Repos
'Microsoft Visual Basic for Applications Extensibility 5.3

Public Sub ExportSourceFiles(destPath As String, Optional IncludeForms As Boolean = False)
    Dim component As VBComponent
    
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        Select Case component.Type
            Case vbext_ct_ClassModule:
                ExportTargetComponent destPath, component
            Case vbext_ct_StdModule:
                ExportTargetComponent destPath, component
            Case vbext_ct_MSForm:
                If IncludeForms = True Then
                    ExportTargetComponent destPath, component
                End If
            Case Else:
                'do nothing
        End Select
    Next
 
End Sub
Private Function ExportTargetComponent(ByVal destPath As String, ByRef component As VBComponent)
    component.Export destPath & component.Name & ToFileExtension(component.Type)
End Function
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        
        Case vbext_ComponentType.vbext_ct_Document
        
        Case Else
            ToFileExtension = vbNullString
    End Select
 
End Function

Private Function IsVBFileType(ByVal strFile As String) As Boolean
    Dim rtnBln As Boolean
    
    If InStr(1, strFile, ".bas") > 1 Then
        rtnBln = True
        GoTo FOUND:
    End If
    
    If InStr(1, strFile, ".cls") > 1 Then
        rtnBln = True
        GoTo FOUND:
    End If
        
    If InStr(1, strFile, ".frm") > 1 Then
        rtnBln = True
        GoTo FOUND:
    End If

FOUND:
    IsVBFileType = rtnBln
End Function

Public Sub ImportSourceFiles(sourcePath As String)
    Dim file As String
    file = Dir(sourcePath)
    
    While (file <> vbNullString)
        If IsVBFileType(file) Then
            Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & file
        End If
        file = Dir
    Wend
End Sub

Public Function RemoveAllModules()
    Dim project As VBProject
    Set project = Application.VBE.ActiveVBProject
     
    Dim comp As VBComponent
    
    For Each comp In project.VBComponents
        If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.VBComponents.Remove comp
        End If
    Next
End Function
