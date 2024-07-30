VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparisonBlockStrategyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Strategies.Contracts"
'@Interface
Option Explicit

Public Function Create( _
       ByVal Document As IWordDocument, _
       ByVal Writer As IDocumentWriter) As IComparisonBlockDecorationStrat
       
End Function

