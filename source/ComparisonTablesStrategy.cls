VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonTablesStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Strategies.Tables"
Option Explicit
Implements IComparisonBlockDecorationStrat

Private Type TTablesOutputResultStrategy
    TableHandler As IDecorationTableHandler
End Type

Private this As TTablesOutputResultStrategy

Public Property Get Self() As IComparisonBlockDecorationStrat
    Set Self = Me
End Property

Public Property Get TableHandler() As IDecorationTableHandler
    Set TableHandler = this.TableHandler
End Property
Public Property Set TableHandler(ByVal RHS As IDecorationTableHandler)
    Set this.TableHandler = RHS
End Property

Public Sub FillBlock( _
    ByVal Model As List_IOutputResultRowModel_, _
    ByVal ParticipantNumber As String)
    
    Dim Table As IWordTable
    Set Table = TableHandler.GetTable(1)
    
    Dim Decorator As ITableRowAppendStrategy
    Set Decorator = New ComparisonTableStrategy

    Dim i As Long
    For i = 1 To Model.Count

        Dim Predicate As Boolean
        Predicate = Model(i).ParticipantNumber = ParticipantNumber
                
        Decorator.AppendRow Table, Model(i), Predicate

    Next
    
    ApplyTableDecoration Decorator, Table
    
End Sub
    Private Sub ApplyTableDecoration( _
            ByVal Decorator As ITableDecorationStrategy, _
            ByVal Table As IWordTable)
            
        Decorator.ApplyTableDecoration Table
    
    End Sub
'---------------------------------------------
Private Sub IComparisonBlockDecorationStrat_FillBlock(ByVal Model As List_IOutputResultRowModel_, ByVal ParticipantNumber As String)
     FillBlock Model, ParticipantNumber
End Sub

