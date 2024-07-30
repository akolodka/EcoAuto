Attribute VB_Name = "RibbonCallbacks"
'@Folder("eco_transferToWord.Ribbon")
Option Explicit

'HelpButton (элемент: button, атрибут: onAction), 2010+
Private Sub ribbon_ShowHelp(Control As IRibbonControl)
    Help
End Sub

'btnTransferToWord (элемент: button, атрибут: onAction), 2010+
Private Sub ribbon_TransferToWord(Control As IRibbonControl)
    Main
End Sub

