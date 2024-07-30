VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Initialization.Data Lists"
'@PredeclaredId
Option Explicit

Public Function CreateStaticFiles() As List_IFile_
    
    Dim Result As List_IFile_
    Set Result = New List_IFile_
    
    Result.Add FileDTO.Create(" ...\firstStaticFile.docx", "firstStaticFile")
    Result.Add FileDTO.Create(" ...\secondStaticFile.docx", "secondStaticFile")
    Result.Add FileDTO.Create(" ...\thirdStaticFile.txt", "thirdStaticFile")
                   
    Set CreateStaticFiles = Result
    
End Function

Public Function CreateSubcontractFiles() As List_IFile_
    
    Dim Result As List_IFile_
    Set Result = New List_IFile_
    
    Result.Add FileDTO.Create(" ...\_default.txt", "_default")
    Result.Add FileDTO.Create(" ...\firstSubcontractFile.txt", "firstSubcontractFile")
    Result.Add FileDTO.Create(" ...\secondSubcontractFile.txt", "secondSubcontractFile")
                   
    Set CreateSubcontractFiles = Result
    
End Function

Public Function CreateDocumentTeplateFiles() As List_IFile_
    
    Dim Result As List_IFile_
    Set Result = New List_IFile_
    
    Result.Add FileDTO.Create(" ...\firstTemplateFile.docx", "firstTemplateFile")
    Result.Add FileDTO.Create(" ...\secondTemplateFile.docx", "secondTemplateFile")
                   
    Set CreateDocumentTeplateFiles = Result
    
End Function

Public Function CreateAuxillaryComparisonFiles() As List_IFile_
    
    Dim Result As List_IFile_
    Set Result = New List_IFile_
    
    Result.Add FileDTO.Create(" ...\measuredValuesTable.docx", "measuredValuesTable")
    Result.Add FileDTO.Create(" ...\referenceValuesTable.docx", "referenceValuesTable")
    Result.Add FileDTO.Create(" ...\comparisonTables.docx", "comparisonTables")
    Result.Add FileDTO.Create(" ...\comparisonCharts.docx", "comparisonCharts")
                   
    Set CreateAuxillaryComparisonFiles = Result
    
End Function

