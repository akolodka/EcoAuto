VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Initialization.File System Object.Content"
'@PredeclaredId
Option Explicit
Implements IFileContent

Private Type TFileContent
    Text As String
End Type

Private this As TFileContent

Public Function Create( _
       ByVal Text As String) As IFileContent

    With New FileContent
        
        .Text = Text
        Set Create = .Self
        
    End With

End Function

Public Property Get Text() As String
    Text = this.Text
End Property
Public Property Let Text(ByVal RHS As String)
    this.Text = GetSafeTextValue(RHS)
End Property

Private Function GetSafeTextValue( _
        ByVal Value As String) As String
        
    If Value = vbNullString Then _
       Exit Function
        
    Dim Result As String
    Result = Trim$(Value)
        
    Do While (Predicate(Result) = True)
        Result = Left$(Result, Len(Result) - 1)
    Loop
        
    GetSafeTextValue = Result
        
End Function

Private Function Predicate( _
        ByVal Value As String) As Boolean
            
    Dim rightChar As String
    rightChar = Right$(Value, 1)
            
    'Пустые символы имеют значения ASCII от 0 до 32
    Const BlankSymAscii As Long = 32
            
    Predicate = (Asc(rightChar) <= BlankSymAscii)
        
End Function

Public Property Get Self() As IFileContent
    Set Self = Me
End Property

'---------------------------------------------
Private Property Get IFileContent_Text() As String
    IFileContent_Text = Text
End Property


