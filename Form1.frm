VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "STOP"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Hello Manu How are you ?"
      Top             =   600
      Width           =   5175
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5640
      Top             =   120
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu MNmerlin 
      Caption         =   "(MNmerlin)"
      Begin VB.Menu MNmerlinHide 
         Caption         =   "Hide - Show"
      End
      Begin VB.Menu MNmerlinTime 
         Caption         =   "What Time is it ?"
      End
      Begin VB.Menu MNmerlinDate 
         Caption         =   "Date is ..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim merlin As IAgentCtlCharacterEx
Dim req As IAgentCtlRequest

Dim month(1 To 12) As String

'==================================================================================
'
'==================================================================================
Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    PopupMenu MNmerlin
End Sub

Private Sub Agent1_DblClick(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    merlin.Speak ("I said you Don't double click me. I don't want you to CLICK ME !")
End Sub

Private Sub Agent1_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    Set req = merlin.Speak("My position is " & CStr(merlin.Left) & ". I don't like this place ! Please move me !")
End Sub


Private Sub Command1_Click()
    merlin.Stop
    merlin.Speak "Why !"
End Sub


Private Sub Form_Load()
    Agent1.Characters.Load "merlin", App.Path & "\Merlin.acs"
    Set merlin = Agent1.Characters("merlin")
    merlin.SoundEffectsOn = True
    merlin.LanguageID = &H409
    With Form1
        merlin.Left = (.Left + .Width \ 2) \ 15
        merlin.Top = (.Top + .Height \ 2) \ 15
    End With
    merlin.Commands.Add "Hello", "Hello ..."
    merlin.Commands.Add "Time", "Time ..."
    merlin.AutoPopupMenu = False
    merlin.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim toto
    merlin.Stop
    merlin.Show
    merlin.MoveTo x \ 15, y \ 15
End Sub


Private Sub MNmerlinDate_Click()
    merlin.Speak CStr(Date)
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MNmerlinTime_Click()
    Dim h As String
    Dim m As String

    h = Mid(CStr(Time), 1, 2)
    m = Mid(CStr(Time), 4, 2)
    merlin.Speak "It is " & h & " hours " & m & " minutes"
End Sub



Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then merlin.Speak Text1
        
End Sub

