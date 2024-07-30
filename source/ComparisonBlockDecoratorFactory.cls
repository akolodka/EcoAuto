VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonBlockDecoratorFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration"
'@PredeclaredId
Option Explicit
Implements IComparisonBlockDecoratorFactor

Private Type TComparisonBlockDecoratorFactory

    DocumentHandler As IDecorationDocumentHandler
    WriterFactory As IDocumentWriterFactory
    Participants As List_IParticipant_
    
End Type

Private this As TComparisonBlockDecoratorFactory

Public Function Create( _
       ByVal Word As IWordDocumentsService, _
       ByVal WriterFactory As IDocumentWriterFactory, _
       ByVal Participants As List_IParticipant_) As IComparisonBlockDecoratorFactor
       
    With New ComparisonBlockDecoratorFactory
    
        Set .DocumentHandler = DecorationDocumentHandler.Create(Word)
        Set .WriterFactory = WriterFactory
        Set .Participants = Participants
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get DocumentHandler() As IDecorationDocumentHandler
    Set DocumentHandler = this.DocumentHandler
End Property
Public Property Set DocumentHandler(ByVal RHS As IDecorationDocumentHandler)
    Set this.DocumentHandler = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get WriterFactory() As IDocumentWriterFactory
    Set WriterFactory = this.WriterFactory
End Property
Public Property Set WriterFactory(ByVal RHS As IDocumentWriterFactory)
    Set this.WriterFactory = RHS
End Property

Public Property Get Self() As IComparisonBlockDecoratorFactor
    Set Self = Me
End Property
'---------------------------------------------
Private Function IComparisonBlockDecoratorFactor_Create( _
        ByVal Strategy As IComparisonBlockStrategyFactory, _
        ByVal ParticipantTemplateFiles As List_IFile_) As IComparisonBlockDecorator
    
    With New ComparisonBlockDecorator
        
        Set .Document = DocumentHandler.GetDocument(Strategy, ParticipantTemplateFiles)
        
        Set .Decorator = Strategy.Create( _
            .Document, WriterFactory.Create(.Document))
        
        Set .Multiplier = ComparisonBlockMultiplier.Create(.Document)
        Set .Participants = Participants
        
        Set .Writer = ComparisonHeaderWriter.Create( _
                WriterFactory.Create(.Document))
    
        Set IComparisonBlockDecoratorFactor_Create = .Self
    
    End With
    
End Function


