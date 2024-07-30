VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelCellAddressParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Reference Values Validator.Address Parser")
'@PredeclaredId
Option Explicit
Implements IExcelCellAddressParser

Private Type TExcelCellAddressParser

    BlockIndex As Long
    ParticipantIndex As Long
    
End Type

Private this As TExcelCellAddressParser

Public Function Create( _
       ByVal BlockIndex As Long, _
       ByVal ParticipantIndex As Long) As IExcelCellAddressParser
    
    With New ExcelCellAddressParser
    
        .BlockIndex = BlockIndex
        .ParticipantIndex = ParticipantIndex
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get BlockIndex() As Long
    BlockIndex = this.BlockIndex
End Property
Public Property Let BlockIndex(ByVal RHS As Long)
    this.BlockIndex = RHS
End Property

Public Property Get ParticipantIndex() As Long
    ParticipantIndex = this.ParticipantIndex
End Property
Public Property Let ParticipantIndex(ByVal RHS As Long)
    this.ParticipantIndex = RHS
End Property

Public Property Get Self() As IExcelCellAddressParser
    Set Self = Me
End Property

Public Function GetRow() As Long
    
    Const IninialRowShift As Long = 3
    
    Dim Result As Long
    Result = IninialRowShift + ParticipantIndex
    
    GetRow = Result
    
End Function
Public Function GetColumn() As Long
    
    Const InitialColumnShift As Long = 6
    Const BlockShift As Long = 6
    Const ColumnBlockShift As Long = 4
    
    Dim Result As Long
    Result = InitialColumnShift + BlockShift * (BlockIndex - 1) + ColumnBlockShift
    
    GetColumn = Result
    
End Function
'---------------------------------------------
Private Function IExcelCellAddressParser_GetRow() As Long
    IExcelCellAddressParser_GetRow = GetRow
End Function
Private Function IExcelCellAddressParser_GetColumn() As Long
    IExcelCellAddressParser_GetColumn = GetColumn
End Function


