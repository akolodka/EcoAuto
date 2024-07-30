VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MeasuredValuesHandlerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates.Values Ranges Filler.Measured Values Range Handler"
'@PredeclaredId
Option Explicit
Implements IMeasuredValuesHandlerFactory

Private Type TMeasuredValuesHandlerFactory
    
    Fso As IFileSystemObject
    WorkbookPath As String
    
End Type

Private this As TMeasuredValuesHandlerFactory

Public Function Create( _
       ByVal Fso As IFileSystemObject, _
       ByVal WorkbookPath As String) As IMeasuredValuesHandlerFactory
    
    With New MeasuredValuesHandlerFactory
        
        Set .Fso = Fso
        .WorkbookPath = WorkbookPath
        
        Set Create = .Self
    
    End With

End Function

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get WorkbookPath() As String
    WorkbookPath = this.WorkbookPath
End Property
Public Property Let WorkbookPath(ByVal RHS As String)
    this.WorkbookPath = RHS
End Property

Public Property Get Self() As IMeasuredValuesHandlerFactory
    Set Self = Me
End Property

Private Function IMeasuredValuesHandlerFactory_Create( _
        ByVal ControlObjectFactoryNumber As String) As IMeasuredValuesHandler
    
    With New MeasuredValuesHandler
        
        Set .Fso = Fso
        .WorkbookPath = WorkbookPath
        
        .ControlObjectFactoryNumber = ControlObjectFactoryNumber
        .Initialize
        
        Set IMeasuredValuesHandlerFactory_Create = .Self
        
    End With
    
End Function


