VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInputValuesHandlerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Transfer Process.Values Handling.User Input"
'@PredeclaredId
Option Explicit
Implements IValuesHandlerFactory

Const SubcontractKey As String = "subcontractDescription"
Const RespondentKey As String = "respondent"

Private Type TUserInputValuesHandlerFactory
    
    Subcontract As IKeyValuePair
    Respondent As IKeyValuePair
    
End Type

Private this As TUserInputValuesHandlerFactory

Public Function Create( _
       ByVal Model As IUserInputTransferModel) As IValuesHandlerFactory

    With New UserInputValuesHandlerFactory
        
        Set .Subcontract = KeyValuePair.Create(SubcontractKey, Model.Subcontract)
        Set .Respondent = KeyValuePair.Create(RespondentKey, Model.Respondent)
        
        Set Create = .Self
        
    End With

End Function

Public Property Get Subcontract() As IKeyValuePair
    Set Subcontract = this.Subcontract
End Property
Public Property Set Subcontract(ByVal RHS As IKeyValuePair)
    Set this.Subcontract = RHS
End Property

Public Property Get Respondent() As IKeyValuePair
    Set Respondent = this.Respondent
End Property
Public Property Set Respondent(ByVal RHS As IKeyValuePair)
    Set this.Respondent = RHS
End Property

Public Property Get Self() As IValuesHandlerFactory
    Set Self = Me
End Property

'---------------------------------------------
Private Function IValuesHandlerFactory_Create( _
        ByVal Writer As IDocumentWriter) As IValuesHandler
    
    With New UserInputValuesHandler
        
        Set .Subcontract = Subcontract
        Set .Respondent = Respondent
        
        Set .Writer = Writer
        
        Set IValuesHandlerFactory_Create = .Self
        
    End With
End Function
