VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParticipantTablesDecoratorFacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Participant Tables Decoration"
'@PredeclaredId
Option Explicit
Implements IParcicipantTablesDecoratorFact

Private Type TComparisonDecorationServiceFactory
    DocumentHandler As IDecorationDocumentHandler
End Type

Private this As TComparisonDecorationServiceFactory

Public Function Create( _
       ByVal Word As IWordDocumentsService) As IParcicipantTablesDecoratorFact
    
    With New ParticipantTablesDecoratorFacto
    
        Set .DocumentHandler = DecorationDocumentHandler.Create(Word)
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Self() As IParcicipantTablesDecoratorFact
    Set Self = Me
End Property

Public Property Get DocumentHandler() As IDecorationDocumentHandler
    Set DocumentHandler = this.DocumentHandler
End Property
Public Property Set DocumentHandler(ByVal RHS As IDecorationDocumentHandler)
    Set this.DocumentHandler = RHS
End Property

Private Function IParcicipantTablesDecoratorFact_Create( _
        ByVal Strategy As IComparisonDecorationStrategy, _
        ByVal ParticipantTableTemplateFiles As List_IFile_, _
        ByVal ParticipantResults As List_IParticipantResult_) As IParticipantTablesDecorator
    
    With New ParticipantTablesDecorator
    
        Set .Strategy = Strategy
        Set .Document = DocumentHandler.GetDocument(Strategy, ParticipantTableTemplateFiles)
        Set .ParticipantResults = ParticipantResults
    
        Set IParcicipantTablesDecoratorFact_Create = .Self
    
    End With
    
End Function


