VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Help.Product Version"
'@PredeclaredId
Option Explicit

Implements IVersionHandler
Implements IDeveloperVersionHandler

Private Const DateAnchorProperty As String = "Comments"
Private Const NumberAnchorProperty As String = "Category"

Private Type TVersionHandler

    NumberValue As String
    DateValue As String
    Version As String
    
End Type

Private this As TVersionHandler

Public Function Create() As IVersionHandler
    
    With New VersionHandler
        
        .ReadFromProperties
        Set Create = .Self
        
    End With
    
End Function

Public Function ReadFromProperties()
    
    DateValue = ThisWorkbook.BuiltinDocumentProperties(DateAnchorProperty)
    NumberValue = ThisWorkbook.BuiltinDocumentProperties(NumberAnchorProperty)
    
    Version = "Версия " & NumberValue & " от " & DateValue & "."
    
End Function

Public Property Get NumberValue() As String
    NumberValue = this.NumberValue
End Property
Public Property Let NumberValue(ByVal RHS As String)
    this.NumberValue = RHS
End Property

Public Property Get DateValue() As String
    DateValue = this.DateValue
End Property
Public Property Let DateValue(ByVal RHS As String)
    this.DateValue = RHS
End Property

Public Property Get Version() As String
    Version = this.Version
End Property
Public Property Let Version(ByVal RHS As String)
    this.Version = RHS
End Property

Public Property Get Self() As IVersionHandler
    Set Self = Me
End Property

'---------------------------------------------
Private Property Get IVersionHandler_Version() As String
    IVersionHandler_Version = Version
End Property

'---------------------------------------------
Private Sub IDeveloperVersionHandler_ApplyNewVersionNumber()
    
    Dim Message As String
    Message = "Текущая метка версии:" & vbNewLine & _
              "" & NumberValue & " от " & DateValue & " " & vbNewLine & _
              "" & vbNewLine & _
              "Новая версия от " & Date & ":"
    
    Dim currentNumber As Double
    currentNumber = Replace(NumberValue, ".", Application.DecimalSeparator)
    
    Dim newNumber As Double
    newNumber = currentNumber + 0.01
    
    Dim suggestedValue As String
    suggestedValue = Replace(newNumber, Application.DecimalSeparator, ".")
    
    Dim newVersion As String
    newVersion = InputBox(Message, , suggestedValue)
    
    If newVersion = vbNullString Then _
       Exit Sub
    
    ThisWorkbook.BuiltinDocumentProperties(NumberAnchorProperty) = newVersion
    ThisWorkbook.BuiltinDocumentProperties(DateAnchorProperty) = Format$(Date, "dd.mm.yyyy")
    
    ThisWorkbook.Save
    
End Sub


