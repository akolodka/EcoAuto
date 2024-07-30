VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "eco_transferToWord.Core.Progress Bar"
Option Explicit

Implements IProgressServicePrimary
Implements IProgressService
Implements IProgressServiceDisposer

Private Type TProgressService

    Color As ProgressColor
    View As ProgressDialog
    
    Step As Double
    ProgressAbsoluteValue As Double
    
    isDialogShowed As Boolean
    
End Type

Public Enum ProgressColor
    
    BlueSea = &HC00000
    GreenGrass = &H8000&
    RedVine = &HC0&
    
End Enum

Private this As TProgressService

Public Function Create( _
       Optional ByVal Color As ProgressColor = ProgressColor.BlueSea) As IProgressServicePrimary
    
    With New ProgressBar
    
        .Color = Color
        Set .View = New ProgressDialog
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Self() As IProgressService
    Set Self = Me
End Property

Public Property Get Color() As ProgressColor
    Color = this.Color
End Property
Public Property Let Color(ByVal RHS As ProgressColor)
    this.Color = RHS
End Property

Public Property Get View() As ProgressDialog
    Set View = this.View
End Property
Public Property Set View(ByVal RHS As ProgressDialog)
    Set this.View = RHS
End Property

Public Property Get Step() As Double
    Step = this.Step
End Property
Public Property Let Step(ByVal RHS As Double)
    this.Step = RHS
End Property

Private Property Get Dialog() As IDialogView
    Set Dialog = View
End Property
'---------------------------------------------
Private Function IProgressServicePrimary_Initialize( _
        ByVal TotalCount As Long) As IProgressService
    
    Step = View.ProgressTargetFrame.Width / TotalCount
    
    View.ProgressValueBlackLabel.Left = View.ProgressValueWhiteLabel.Left

    With View.ProgressAchievedFrame

        .Width = 0.0001
        .BackColor = Color

    End With
    
    Set IProgressServicePrimary_Initialize = Self
    
End Function
'---------------------------------------------
Private Sub IProgressService_Show( _
        ByVal Message As String, _
        Optional ByVal Description As String = vbNullString)
    
    View.Caption = IIf(Message <> vbNullString, _
                       Message, "Выполнение...")
        
    View.DescriptionLabel = Description
    
    Const HeightMinimun As Single = 52.5
    Const HeightMaximum As Long = 69
    
    View.Height = IIf(Description = vbNullString, _
                      HeightMinimun, HeightMaximum)

    this.ProgressAbsoluteValue = this.ProgressAbsoluteValue + Step
    View.ProgressAchievedFrame.Width = this.ProgressAbsoluteValue
     
    Dim progressRelativeValue As String
    progressRelativeValue = Format$( _
                            this.ProgressAbsoluteValue / View.ProgressTargetFrame.Width, _
                            "0 %")
    
    View.ProgressValueWhiteLabel.Caption = progressRelativeValue
    View.ProgressValueBlackLabel.Caption = progressRelativeValue
    
    View.Repaint
    
    If this.isDialogShowed = True Then _
       Exit Sub
    
    Dialog.Show
    this.isDialogShowed = True

End Sub
'---------------------------------------------
Private Sub IProgressServiceDisposer_Dispose()

    Dialog.Hide
    this.isDialogShowed = False
    
End Sub


