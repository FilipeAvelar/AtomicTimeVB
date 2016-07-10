Attribute VB_Name = "modMain"
Option Explicit

'Change to your closest Internet Time Server
'Public Const TIME_SERVER As String = "ntps1-0.cs.tu-berlin.de" ' "Rolex.PeachNet.edu"
Public TIME_SERVER As String '= "ntps1-0.cs.tu-berlin.de" ' "Rolex.PeachNet.edu"
Public ServerIndex As Integer
Public BatchMode As Boolean

Sub Main()

    
End Sub

Public Function dtString() As String
    'Quick Date and Time string w/full 4 digit year
    dtString = Format(Now, "dd/mm/yyyy hh:mm:ss") 'AM/PM -
End Function

'Public Sub SetTime(sTimeServer As String)
'    Dim oInetTime As cInetTime
'    Set oInetTime = New cInetTime
'
'    oInetTime.TimeServer = sTimeServer
'    oInetTime.SetTime
'
'    Set oInetTime = Nothing
'
'End Sub

