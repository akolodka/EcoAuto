Attribute VB_Name = "RibbonCallbacks"
'@Folder("eco_transferToWord.Ribbon")
Option Explicit

'HelpButton (�������: button, �������: onAction), 2010+
Private Sub ribbon_ShowHelp(Control As IRibbonControl)
    Help
End Sub

'btnTransferToWord (�������: button, �������: onAction), 2010+
Private Sub ribbon_TransferToWord(Control As IRibbonControl)
    Main
End Sub

