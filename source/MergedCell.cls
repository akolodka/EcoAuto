VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MergedCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("eco_transferToWord.Core.Word.Table.Merged Cell")
Option Explicit
Implements IMergedCell

Private Type TMergedCell

    StartRow As Long
    EndRow As Long
    Value As String
    
End Type

Private this As TMergedCell

Public Function Create( _
       ByVal RowIndex As Long, _
       ByVal Value As String) As IMergedCell
    
    With New MergedCell
    
        .StartRow = RowIndex
        .EndRow = RowIndex
        .Value = Value
        
        Set Create = .Self
    
    End With

End Function

Public Property Get Self() As IMergedCell
    Set Self = Me
End Property

Public Property Get StartRow() As Long
    StartRow = this.StartRow
End Property
Public Property Let StartRow(ByVal RHS As Long)
    this.StartRow = RHS
End Property

Public Property Get EndRow() As Long
    EndRow = this.EndRow
End Property
Public Property Let EndRow(ByVal RHS As Long)
    this.EndRow = RHS
End Property

Public Property Get Value() As String
    Value = this.Value
End Property
Public Property Let Value(ByVal RHS As String)
    this.Value = RHS
End Property

'---------------------------------------------
Private Property Get IMergedCell_StartRow() As Long
    IMergedCell_StartRow = StartRow
End Property

Private Property Get IMergedCell_EndRow() As Long
    IMergedCell_EndRow = EndRow
End Property
Private Property Let IMergedCell_EndRow(ByVal RHS As Long)
    EndRow = RHS
End Property

Private Property Get IMergedCell_Value() As String
    IMergedCell_Value = Value
End Property


