VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferMenuCacheHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Cache Handler"
'@PredeclaredId
Option Explicit
Implements ITransferMenuCacheHandler

Const SelectedSubcontractCacheKey As String = "selectedSubcontract"
Const ControlObjectFactoryNumberCacheKey As String = "factoryNumber"
Const RespondentCacheKey As String = "respondent"

Private Type TTransferMenuCacheHandler

    Fso As IFileSystemObject
    CacheIO As ICacheIOService
    Editor As ICacheContentEditService
    
    Respondent As String
    FactoryNumber As String
    SelectedSubcontract As List_String_
    
    Parser As SelectedSubcontractSerializer
    
End Type

Private this As TTransferMenuCacheHandler

Public Function Create( _
       ByVal FileSystemObject As IFileSystemObject) As ITransferMenuCacheHandler
    
    With New TransferMenuCacheHandler
        
        Set .Fso = FileSystemObject
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get CacheIO() As ICacheIOService
    Set CacheIO = this.CacheIO
End Property
Public Property Set CacheIO(ByVal RHS As ICacheIOService)
    Set this.CacheIO = RHS
End Property

Public Property Get Editor() As ICacheContentEditService
    Set Editor = this.Editor
End Property
Public Property Set Editor(ByVal RHS As ICacheContentEditService)
    Set this.Editor = RHS
End Property

Public Property Get Respondent() As String
    Respondent = this.Respondent
End Property
Public Property Let Respondent(ByVal RHS As String)
    this.Respondent = RHS
End Property

Public Property Get ControlObjectFactoryNumber() As String
    ControlObjectFactoryNumber = this.FactoryNumber
End Property
Public Property Let ControlObjectFactoryNumber(ByVal RHS As String)
    this.FactoryNumber = RHS
End Property

Public Property Get SelectedSubcontract() As List_String_
    Set SelectedSubcontract = this.SelectedSubcontract
End Property
Public Property Set SelectedSubcontract(ByVal RHS As List_String_)
    Set this.SelectedSubcontract = RHS
End Property

Public Property Get Parser() As SelectedSubcontractSerializer
    Set Parser = this.Parser
End Property
Public Property Set Parser(ByVal RHS As SelectedSubcontractSerializer)
    Set this.Parser = RHS
End Property

Public Property Get Self() As ITransferMenuCacheHandler
    
    Initialize
    Set Self = Me
    
End Property

Private Sub Initialize()
        
    Set CacheIO = CacheIOService.Create(Fso)
    
    Dim Content As IFileContent
    Set Content = CacheIO.ReadContent
        
    Set Editor = CacheContentEditService.Create(Content)
        
    Respondent = Editor.GetFromContent(RespondentCacheKey)
    ControlObjectFactoryNumber = Editor.GetFromContent(ControlObjectFactoryNumberCacheKey)
        
    Dim selectedValueAsString As String
    selectedValueAsString = Editor.GetFromContent(SelectedSubcontractCacheKey)
        
    Set Parser = New SelectedSubcontractSerializer
    Set SelectedSubcontract = Parser.DeserializeIntoList(selectedValueAsString)
        
End Sub

'---------------------------------------------
Private Property Get ITransferMenuCacheHandler_Respondent() As String
    ITransferMenuCacheHandler_Respondent = Respondent
End Property
Private Property Let ITransferMenuCacheHandler_Respondent(ByVal Value As String)
    Respondent = Value
End Property

Private Property Get ITransferMenuCacheHandler_ControlObjectFactoryNumber() As String
    ITransferMenuCacheHandler_ControlObjectFactoryNumber = ControlObjectFactoryNumber
End Property
Private Property Let ITransferMenuCacheHandler_ControlObjectFactoryNumber(ByVal Value As String)
    ControlObjectFactoryNumber = Value
End Property

Private Property Get ITransferMenuCacheHandler_SelectedSubcontract() As List_String_
    Set ITransferMenuCacheHandler_SelectedSubcontract = SelectedSubcontract
End Property
Private Property Set ITransferMenuCacheHandler_SelectedSubcontract(ByVal Value As List_String_)
    Set SelectedSubcontract = Value
End Property

Private Sub ITransferMenuCacheHandler_Save()
    
    Dim SerializedSubcotractIndexes As String
    SerializedSubcotractIndexes = Parser.ToStringFromList(SelectedSubcontract)
    
    Editor.InsertIntoContent SelectedSubcontractCacheKey, SerializedSubcotractIndexes
    Editor.InsertIntoContent ControlObjectFactoryNumberCacheKey, ControlObjectFactoryNumber
    Editor.InsertIntoContent RespondentCacheKey, Respondent
    
    CacheIO.WriteContent Editor.Content
    
End Sub


