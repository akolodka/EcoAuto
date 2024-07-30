VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuesRangeFillerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates.Values Ranges Filler"
'@PredeclaredId
Option Explicit
Implements IValuesRangeFillerFactory

Private Type TValuesRangeFillerFactory
    HandlerFactory As IMeasuredValuesHandlerFactory
End Type

Private this As TValuesRangeFillerFactory

Public Function Create( _
       ByVal Fso As IFileSystemObject, _
       ByVal WorkbookPath As String) As IValuesRangeFillerFactory
    
    With New ValuesRangeFillerFactory
        
        Set .HandlerFactory = MeasuredValuesHandlerFactory.Create(Fso, WorkbookPath)
        Set Create = .Self
    
    End With

End Function

Public Property Get HandlerFactory() As IMeasuredValuesHandlerFactory
    Set HandlerFactory = this.HandlerFactory
End Property
Public Property Set HandlerFactory(ByVal RHS As IMeasuredValuesHandlerFactory)
    Set this.HandlerFactory = RHS
End Property

Public Property Get Self() As IValuesRangeFillerFactory
    Set Self = Me
End Property
'---------------------------------------------
Private Function IValuesRangeFillerFactory_Create( _
       ByVal ControlObjectFactoryNumber As String, _
       ByVal Participants As List_IParticipant_) As IParticipantValuesRangeFiller

    With New ParticipantValuesRangeFiller
        
        Set .ValuesHandler = HandlerFactory.Create(ControlObjectFactoryNumber)
        Set .Participants = Participants
        
        Set IValuesRangeFillerFactory_Create = .Self
            
    End With

End Function


