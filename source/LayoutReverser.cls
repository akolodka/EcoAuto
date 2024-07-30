VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LayoutReverser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Dialogs.Main Transfer.Presenter Delegate.Layout Reverser")
'@PredeclaredId
Option Explicit
Implements ILayoutReverser
        
Const EnglishKeys As String = "qwertyuiop[]asdfghjkl;'zxcvbnm,."
Const RussianKeys As String = "éöóêåíãøùçõúôûâàïðîëäæýÿ÷ñìèòüáþ"

Private Type TLayoutReverser
    Text As String
End Type

Private this As TLayoutReverser

Public Function Create(ByVal Text As String) As ILayoutReverser
    
    With New LayoutReverser
        
        .Text = Text
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Text() As String
    Text = this.Text
End Property
Public Property Let Text(ByVal RHS As String)
    this.Text = RHS
End Property

Public Property Get Self() As ILayoutReverser
    Set Self = Me
End Property

Public Function GetReversedString() As String
    
    Dim Result As String
    Result = GetLayoutReversedString(EnglishKeys, RussianKeys)
    
    If Result = Text Then _
        Result = GetLayoutReversedString(RussianKeys, EnglishKeys)
    
    GetReversedString = Result
    
End Function
    Private Function GetLayoutReversedString( _
            ByVal FromLayoutKeys As String, _
            ByVal ToLayoutKeys As String) As String
            
        Dim Result As String
            
        Dim i As Long
        For i = 1 To Len(Text)
                
            Dim Sym As String
            Sym = Mid(Text, i, 1)
                
            Dim InstrPosition As Long
            InstrPosition = InStr(FromLayoutKeys, Sym)
                
            If InstrPosition > vbEmpty Then _
               Sym = Mid(ToLayoutKeys, InstrPosition, 1)
            
            Result = Result & Sym
            
        Next
        
        GetLayoutReversedString = Result
            
    End Function
'---------------------------------------------
Private Function ILayoutReverser_GetReversedString() As String
    ILayoutReverser_GetReversedString = GetReversedString
End Function
