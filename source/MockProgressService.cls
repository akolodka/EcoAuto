VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockProgressService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Progress Notification"
Option Explicit

Implements IProgressServicePrimary
Implements IProgressService
Implements IProgressServiceDisposer

Private Type TMockProgressPresenter

    TargetCount As Long
    CurrentCount As Long
    
    TimeStamp As Double
    
End Type

Private this As TMockProgressPresenter

Public Property Get Self() As IProgressService
    Set Self = Me
End Property

Public Property Get TargetCount() As Long
    TargetCount = this.TargetCount
End Property
Public Property Let TargetCount(ByVal RHS As Long)
    this.TargetCount = RHS
End Property

Public Property Get TimeStamp() As Double
    TimeStamp = this.TimeStamp
End Property
Public Property Let TimeStamp(ByVal RHS As Double)
    this.TimeStamp = RHS
End Property

Public Property Get CurrentCount() As Long
    CurrentCount = this.CurrentCount
End Property
Public Property Let CurrentCount(ByVal RHS As Long)
    this.CurrentCount = RHS
End Property
'---------------------------------------------
Private Function IProgressServicePrimary_Initialize( _
    ByVal TotalCount As Long) As IProgressService
    
    TargetCount = TotalCount
    
    Set IProgressServicePrimary_Initialize = Self
    
End Function
'---------------------------------------------
Private Sub IProgressService_Show( _
        ByVal Message As String, _
        Optional ByVal Footer As String = vbNullString)
    
    If TimeStamp = vbEmpty Then _
       TimeStamp = Timer
    
    CurrentCount = CurrentCount + 1
    
Debug.Print
Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Debug.Print "Progress: "; CurrentCount; " из "; TargetCount
Debug.Print "    Message: "; Message
    
    If Footer <> vbNullString Then _
Debug.Print "    "; Footer
        
Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    
End Sub



Private Sub IProgressServiceDisposer_Dispose()

Debug.Print
Debug.Print "Time elapsed: "; Timer - TimeStamp; " s"

End Sub

