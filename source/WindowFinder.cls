VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Results Presentator.Explorer Folder Open.Window Finder"
'@PredeclaredId
Option Explicit
Implements IWindowFinder

Private Type TWindowFinder

    WindowName As String
    IsFinded As Boolean
    
End Type

Private this As TWindowFinder

Public Function Create( _
       ByVal WindowName As String) As IWindowFinder
    
    With New WindowFinder
    
        .WindowName = WindowName
        Set Create = .Self
    
    End With
    
End Function

Public Property Get WindowName() As String
    WindowName = this.WindowName
End Property
Public Property Let WindowName(ByVal RHS As String)
    this.WindowName = RHS
End Property

Public Property Get Self() As IWindowFinder
    Set Self = Me
End Property

Private Function IWindowFinder_GetWindow() As Object
  
    Dim shellObject As New Shell32.Shell
    Dim Window As Object
    
    For Each Window In shellObject.Windows
    
        If (Predicate(Window.locationName) = True) Then
        
            Set IWindowFinder_GetWindow = Window
            this.IsFinded = True
            
            Exit Function
        
        End If
                
    Next

End Function

Private Function Predicate( _
        ByVal locationName As String) As Boolean
    
    Predicate = (locationName = WindowName)
    
End Function

'---------------------------------------------
Private Property Get IWindowFinder_IsWindowFinded() As Boolean
    IWindowFinder_IsWindowFinded = this.IsFinded
End Property


