VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultInterpretationHandlerFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Result Interpretation Handler")
'@PredeclaredId
Option Explicit
Implements IResultInterpretationHandlerFac

Private Type TResultInterpretationHandlerFact

    Strategies As List_IResultInterpretationStra_
    Progress As IProgressService
    
End Type

Private this As TResultInterpretationHandlerFact

Public Function Create( _
       ByVal Strategies As List_IResultInterpretationStra_, _
       ByVal Progress As IProgressService) As IResultInterpretationHandlerFac
    
    With New ResultInterpretationHandlerFact
        
        Set .Strategies = Strategies
        Set .Progress = Progress
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Strategies() As List_IResultInterpretationStra_
    Set Strategies = this.Strategies
End Property
Public Property Set Strategies(ByVal RHS As List_IResultInterpretationStra_)
    Set this.Strategies = RHS
End Property

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get Self() As IResultInterpretationHandlerFac
    Set Self = Me
End Property
'---------------------------------------------
Private Function IResultInterpretationHandlerFac_Create( _
        ByVal Document As IWordDocument, _
        ByVal Results As List_IParticipantResult_) As IResultInterpretationHandler
    
    With New ResultInterpretationHandler
    
        Set .Strategies = Strategies
        Set .Document = Document
        
        Set .Progress = Progress
        
        Set .ParticipantResults = Results
        Set IResultInterpretationHandlerFac_Create = .Self
    
    End With
    
End Function



