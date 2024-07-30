VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferHelpPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Help"
'@PredeclaredId
Option Explicit
Implements ITransferDialogService

Private WithEvents HelpView As TransferHelpDialog
Attribute HelpView.VB_VarHelpID = -1

Private Type TTransferHelpPresenter
    CallBackMenu As ITransferDialogService
End Type

Private this As TTransferHelpPresenter

Public Function Create( _
       Optional ByVal CallBackMenu As ITransferDialogService) As ITransferDialogService

    With New TransferHelpPresenter
        
        Set .View = New TransferHelpDialog
        Set .CallBackMenu = CallBackMenu
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get View() As TransferHelpDialog
    Set View = HelpView
End Property
Public Property Set View(ByVal RHS As TransferHelpDialog)
    Set HelpView = RHS
End Property

Public Property Get CallBackMenu() As ITransferDialogService
    Set CallBackMenu = this.CallBackMenu
End Property
Public Property Set CallBackMenu(ByVal RHS As ITransferDialogService)
    Set this.CallBackMenu = RHS
End Property

Public Property Get Self() As ITransferDialogService
    Set Self = Me
End Property

'---------------------------------------------
Private Sub ITransferDialogService_Show()
    
    View.TextLabel = GetHelpMessage
    View.VersionLabel = GetVersionMessage
    
    Dialog.Show
    
End Sub

Private Function GetHelpMessage() As String
                
    Dim Message As String
    Message = "1 Выберите из списка «Участники тура» позиции для переноса." & _
              " По умолчанию в список попадают строки участников, для которых присвоен номер и указана организация." & _
              " Используйте фильтр при необходимости." & vbNewLine & _
              vbNewLine & _
              "2 Если из списка «Работы по субподряду»" & vbNewLine & _
              "не выбран ни один элемент, при переносе" & vbNewLine & _
              "будет использовано содержимое файла «_default.txt»." & vbNewLine & _
              vbNewLine & _
              "3 Элементы управления формы отобразят подсказки при наведении кусора." & vbNewLine & _
              vbNewLine & _
              "4 Нажмите ESC для выхода из меню." & vbNewLine & _
              vbNewLine & _
              "5 Нажмите Enter для начала переноса в Word."
        
    GetHelpMessage = Message
        
End Function

Private Function GetVersionMessage() As String
        
    Dim Handler As IVersionHandler
    Set Handler = VersionHandler.Create()
        
    Dim Result As String
    Result = Handler.Version
        
    GetVersionMessage = Result
        
End Function

Private Property Get Dialog() As IDialogView
    Set Dialog = View
End Property

'---------------------------------------------
Private Sub HelpView_CloseButtonPressed()
    
    Dialog.Hide
    
    If (IsCallBackMenu = True) Then _
       CallBackMenu.Show
    
End Sub

Private Property Get IsCallBackMenu() As Boolean
    IsCallBackMenu = (CallBackMenu Is Nothing = False)
End Property


