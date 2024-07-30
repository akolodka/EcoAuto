VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExplorerFolderOpener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Results Presentator.Explorer Folder Open"
'@PredeclaredId
Option Explicit
Implements IExplorerFolderOpener

Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Type TExplorerPresentationService
    Fso As IFileSystemObject
End Type

Private this As TExplorerPresentationService

Public Function Create( _
       ByVal Fso As IFileSystemObject) As IExplorerFolderOpener
    
    With New ExplorerFolderOpener
    
        Set .Fso = Fso
        Set Create = .Self
        
    End With

End Function

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Self() As IExplorerFolderOpener
    Set Self = Me
End Property

Private Sub ShowInExplorer( _
        ByVal folderPath As String)

    If (IsCorrect(folderPath) = False) Then _
       Exit Sub
    
    Dim WindowName As String
    WindowName = GetWindowName(folderPath)
    
    Dim Finder As IWindowFinder
    Set Finder = WindowFinder.Create(WindowName)
    
    Dim Window As Object
    Set Window = Finder.GetWindow()
    
    If (Finder.IsWindowFinded = True) Then
        ActivateWindow Window
    Else
        
        If (Fso.FolderExists(folderPath) = True) Then _
           Fso.OpenByShell folderPath
            
    End If

End Sub

Private Function IsCorrect( _
        ByVal Path As String) As Boolean
        
    IsCorrect = (Path <> Application.PathSeparator)
        
End Function

Private Function GetWindowName( _
        ByVal Path As String) As String
        
    Dim Result As String
    Result = Fso.GetBaseName(Path)
        
    If Result = "Desktop" Then _
       Result = "Рабочий стол"
        
    GetWindowName = Result
        
End Function

Private Sub ActivateWindow( _
        ByVal Window As Object)
    
    Window.Visible = False
    Window.Visible = True

    Const SW_RESTORE As Long = 9

    If (CBool(IsIconic(Window.hwnd)) = True) Then _
       ShowWindow Window.hwnd, SW_RESTORE
        
End Sub

'---------------------------------------------
Private Sub IExplorerFolderOpener_ShowInExplorer(ByVal folderPath As String)
    ShowInExplorer folderPath
End Sub


