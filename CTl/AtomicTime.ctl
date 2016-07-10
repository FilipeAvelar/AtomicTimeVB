VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl AtomicTime 
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   InvisibleAtRuntime=   -1  'True
   Picture         =   "AtomicTime.ctx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   4680
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "AtomicTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function UpdateTime() As Boolean
Dim Echeck As Boolean
    Dim oInetTime As cInetTime
    Set oInetTime = New cInetTime
    
    On Error GoTo ERRO
    
    With oInetTime
        .TimeServer = "200.20.186.75"
        .SetTime Winsock1
        If .ErrorCheck = True And .ReturnedDate <> #12:00:00 AM# Then
            UpdateTime = False
        Else
            UpdateTime = True
            Echeck = False
            On Error Resume Next
            Date = Format(.ReturnedDate, "short date")
            Time = Format(.ReturnedDate, "hh:mm:ss")
        End If
        
    End With
    
    
    Set oInetTime = Nothing
    
    If BatchMode = True And Echeck = False Then Unload Me
    
    BatchMode = False
ERRO:
End Function

Private Sub UserControl_Resize()
    UserControl.Width = 600
    UserControl.Height = 600
End Sub
