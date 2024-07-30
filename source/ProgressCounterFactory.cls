VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressCounterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("eco_transferToWord.Core.Initialization.Progress Counter.Factory")
Option Explicit
Implements IProgressCounterFactory

Private Type TProgressCounterFactory

    StaticValueFilesCount As Long
    ResultDocumentTemplatesCount As Long
    AuxillaryTablesFilesCount As Long
    
End Type

Private this As TProgressCounterFactory

Public Function Create( _
       ByVal StaticValueFilesCount As Long, _
       ByVal AuxillaryTablesFilesCount As Long, _
       ByVal ResultDocumentTemplatesCount As Long) As IProgressCounterFactory
    
    With New ProgressCounterFactory
        
        .StaticValueFilesCount = StaticValueFilesCount
        .AuxillaryTablesFilesCount = AuxillaryTablesFilesCount
        .ResultDocumentTemplatesCount = ResultDocumentTemplatesCount
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get StaticValueFilesCount() As Long
    StaticValueFilesCount = this.StaticValueFilesCount
End Property
Public Property Let StaticValueFilesCount(ByVal RHS As Long)
    this.StaticValueFilesCount = RHS
End Property

Public Property Get ResultDocumentTemplatesCount() As Long
    ResultDocumentTemplatesCount = this.ResultDocumentTemplatesCount
End Property
Public Property Let ResultDocumentTemplatesCount(ByVal RHS As Long)
    this.ResultDocumentTemplatesCount = RHS
End Property

Public Property Get AuxillaryTablesFilesCount() As Long
    AuxillaryTablesFilesCount = this.AuxillaryTablesFilesCount
End Property
Public Property Let AuxillaryTablesFilesCount(ByVal RHS As Long)
    this.AuxillaryTablesFilesCount = RHS
End Property

Public Property Get Self() As IProgressCounterFactory
    Set Self = Me
End Property

'---------------------------------------------
Private Function IProgressCounterFactory_Create( _
    ByVal ParticipantsCount As Long, _
    ByVal ComparisonBlockCount As Long) As IProgressCounter
    
    With New ProgressCounter
    
        .StaticValueFilesCount = StaticValueFilesCount
        .AuxillaryTablesFilesCount = AuxillaryTablesFilesCount
        .ResultDocumentTemplatesCount = ResultDocumentTemplatesCount
        
        .ParticipantsCount = ParticipantsCount
        .ComparisonBlockCount = ComparisonBlockCount
        
        Set IProgressCounterFactory_Create = .Self
        
    End With

End Function

