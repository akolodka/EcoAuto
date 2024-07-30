VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartObjectWriteStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Document IO.IO Strategies")
Option Explicit

Implements ITransferValueIOStrategy
Implements ITransferValueWriter

Public Function IsMatchesForWrite( _
       ByVal Value As ITransferValue) As Boolean
    
    If (Value Is Nothing = True) Then _
       Exit Function
    
    If (IsObject(Value.Value) = False) Then _
       Exit Function
        
    If (Value.Value Is Nothing = True) Then _
       Exit Function
        
    IsMatchesForWrite = TypeOf Value.Value Is IExcelChart
    
End Function
Public Sub InsertIntoDocument( _
       ByVal Value As ITransferValue, _
       ByVal Destination As IEditableWordRange)
    
    If (IsMatchesForWrite(Value) = False) Then _
       Exit Sub
    
    If (Destination Is Nothing = True) Then _
       Exit Sub
    
    Dim Source As IExcelChart
    Set Source = Value.Value
        
    If Source Is Nothing Then _
       Exit Sub
    
    Source.CopyAsPicture
    Destination.PasteAndFormat
    
End Sub
'---------------------------------------------
Private Function ITransferValueIOStrategy_IsMatchesForRead(ByVal FilePath As String) As Boolean
    ITransferValueIOStrategy_IsMatchesForRead = False
End Function
Private Function ITransferValueIOStrategy_IsMatchesForWrite(ByVal Value As ITransferValue) As Boolean
    ITransferValueIOStrategy_IsMatchesForWrite = IsMatchesForWrite(Value)
End Function
'---------------------------------------------
Private Sub ITransferValueWriter_InsertIntoDocument(ByVal Value As ITransferValue, ByVal Destination As IWordRange)
    InsertIntoDocument Value, Destination
End Sub

