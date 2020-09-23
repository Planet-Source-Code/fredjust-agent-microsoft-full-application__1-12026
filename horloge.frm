VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRMmerlin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MERLIN"
   ClientHeight    =   1260
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2280
   Icon            =   "horloge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   2280
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1380
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   0
      Top             =   120
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   480
      Top             =   120
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu MNmerlin 
      Caption         =   "(MNmerlin)"
      Begin VB.Menu MNmerlinAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MNmerlinSep11 
         Caption         =   "-"
      End
      Begin VB.Menu MNmerlinConfig 
         Caption         =   "Config"
      End
      Begin VB.Menu MNmerlinSep00 
         Caption         =   "-"
      End
      Begin VB.Menu MNmerlinHide 
         Caption         =   "Hide - Show"
      End
      Begin VB.Menu MNmerlinTime 
         Caption         =   "What Time is it ?"
      End
      Begin VB.Menu MNmerlinFreq 
         Caption         =   "Freqency"
         Begin VB.Menu MNmerlinFreqLong 
            Caption         =   "Never !"
            Index           =   0
         End
         Begin VB.Menu MNmerlinFreqLong 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MNmerlinFreqLong 
            Caption         =   "Each 5 minutes"
            Index           =   5
         End
         Begin VB.Menu MNmerlinFreqLong 
            Caption         =   "Each 10 minutes"
            Index           =   10
         End
         Begin VB.Menu MNmerlinFreqLong 
            Caption         =   "Each 15 minutes"
            Checked         =   -1  'True
            Index           =   15
         End
         Begin VB.Menu MNmerlinFreqLong 
            Caption         =   "Each 30 minutes"
            Index           =   30
         End
         Begin VB.Menu MNmerlinFreqLong 
            Caption         =   "Each Hours"
            Index           =   60
         End
      End
      Begin VB.Menu MNmerlinSize 
         Caption         =   "Size"
         Begin VB.Menu MNmerlinSizeSmall 
            Caption         =   "Small"
         End
         Begin VB.Menu MNmerlinSizeNormal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu MNmerlinSizeBig 
            Caption         =   "Big"
         End
      End
      Begin VB.Menu MNmerlinSep01 
         Caption         =   "-"
      End
      Begin VB.Menu MNmerlinAnnule 
         Caption         =   "Cancel"
      End
      Begin VB.Menu MNmerlinSep03 
         Caption         =   "-"
      End
      Begin VB.Menu MNmerlinExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FRMmerlin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================
'   Réalisation de Frédéric Just
'   Commentaires remarques et critiques :
'
'   adresse en cours    : fred.just@free.fr
'   site actuel         : http://www.fredjust.com
'   adresse de secours  : fredjust@hotmail.com
'==================================================================================


Option Explicit

Dim fini, deb
Dim fso As FileSystemObject

Dim m As String

'==================================================================================
'
'==================================================================================
Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
    If FRMbulleVisible Then FRMbulle.SetFocus
    If FRMmessageVisible Then FRMmessage.SetFocus
    If Button = 4098 Or Button = 2 Then
        MNmerlinHide.Caption = IIf(merlin.Visible, "Hide", "Show")
        PopupMenu MNmerlin
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Agent1_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    lasttop = merlin.Top
    lastleft = merlin.Left
    merlin.Stop
    If FRMbulleVisible Then MNmerlinConfig_Click
    If FRMmessageVisible Then AfficheMessage
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Agent1_RequestComplete(ByVal Request As Object)
    If Request = fini Then
        Unload Me
    End If
    If Not InMem Then
        Load FRMbulle
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Agent1_Show(ByVal CharacterID As String, ByVal Cause As Integer)
    merlin.Top = lasttop
    merlin.Left = lastleft
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Load()
    On Error GoTo gestion_erreur
    
    If App.PrevInstance Then
        MsgBox "TimeAgent is in memory yet !", vbCritical, "Time Agent Error"
        End
    End If
    
    ShowAbout = False
    
    Set fso = New FileSystemObject
    
    InMem = False
    
    SaveAs = GetSetting("TimeAgent", "option", "saveas", App.Path & "\Messages.lv")
    
    NextMessageDate = #12/12/2099#
    
    chaineagent = GetSetting("TimeAgent", "string", "Agent", fso.GetSpecialFolder(WindowsFolder) & "\msagent\chars" & "\Genie.acs")
    If Not fso.FileExists(chaineagent) Then
        MsgBox "Sorry but I don't find any agent, Please Tell where are they ...", vbInformation
        GetOtherAgent
    End If
    
    If Not fso.FileExists(chaineagent) Then
        MsgBox "Sorry,But this program need an Agent, Retry later ...", vbInformation
        End
    End If
    
    Agent1.Characters.Load "merlin", chaineagent
    Set merlin = Agent1.Characters("merlin")
    merlin.SoundEffectsOn = True
    merlin.LanguageID = &H409

    lastleft = Screen.Width \ 15 - merlin.Width - 20
    lasttop = Screen.Height \ 15 - merlin.Height - 20
    merlin.AutoPopupMenu = False
    merlin.Show

    Frequence = CLng(GetSetting("TimeAgent", "Option", "Frequence", "30"))
    ChaineTime = GetSetting("TimeAgent", "string", "TimeFormat", "It is #H hours, #M minutes.")
    ChaineBonjour = GetSetting("TimeAgent", "string", "OnShow", "Yes Master !")
    ChaineFin = GetSetting("TimeAgent", "string", "OnExit", "Your desire is an order.")
    
    If ChaineBonjour <> "" Then
        Set deb = merlin.Speak(ChaineBonjour)
    End If

    Exit Sub
gestion_erreur:
    MsgBox Err.Description, vbCritical, "ERREUR n°" & CStr(Err.Number)
    End
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Unload(Cancel As Integer)
    Unload FRMbulle
    Unload FRMmessage
    Set FRMmessage = Nothing
    Set FRMbulle = Nothing
    Set merlin = Nothing
    Set fso = Nothing
    SaveSetting "TimeAgent", "string", "TimeFormat", ChaineTime
    SaveSetting "TimeAgent", "string", "OnShow", ChaineBonjour
    SaveSetting "TimeAgent", "string", "OnExit", ChaineFin
    SaveSetting "TimeAgent", "Option", "Frequence", CStr(Frequence)
    SaveSetting "TimeAgent", "Option", "saveas", SaveAs
End Sub

'==================================================================================
'
'==================================================================================
Private Function StyleEnFonction(ByVal X As Long, ByVal Y As Long) As Long
    StyleEnFonction = 0
    If X < Screen.Width / 30 Then StyleEnFonction = StyleEnFonction + 1
    If Y < Screen.Height / 30 Then StyleEnFonction = 5 - StyleEnFonction
End Function




Private Sub MNday_Click(index As Integer)
End Sub

Private Sub MNhour_Click(index As Integer)
End Sub

Private Sub MNmerlinAbout_Click()
    ShowAbout = True
    AfficheMessage
    
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MNmerlinConfig_Click()
    Dim LeStyle As Long
    
    If Not FRMbulle.Visible Then LisContenuLISTVIEWdepuisFichier FRMbulle.LV, SaveAs
    
    If merlin.Visible Then
        LeStyle = StyleEnFonction(lastleft + merlin.Width / 2, lasttop + merlin.Height / 2)
        If LeStyle = 0 Then
            FRMbulle.Top = lasttop * Screen.TwipsPerPixelY - FRMbulle.Height + 150
            FRMbulle.Left = lastleft * Screen.TwipsPerPixelX - FRMbulle.Width + merlin.Width * 5
        End If
        If LeStyle = 1 Then
            FRMbulle.Top = lasttop * Screen.TwipsPerPixelY - FRMbulle.Height + 150
            FRMbulle.Left = lastleft * Screen.TwipsPerPixelX + merlin.Width * 10
        End If

        If LeStyle = 4 Then
            FRMbulle.Top = lasttop * Screen.TwipsPerPixelY + merlin.Height * 15
            FRMbulle.Left = lastleft * Screen.TwipsPerPixelX + merlin.Width * 10
        End If

        If LeStyle = 5 Then
            FRMbulle.Top = lasttop * Screen.TwipsPerPixelY + merlin.Height * 15
            FRMbulle.Left = lastleft * Screen.TwipsPerPixelX - FRMbulle.Width + merlin.Width * 5
        End If
    Else
        FRMbulle.Left = (Screen.Width - FRMbulle.Width) / 2
        FRMbulle.Top = (Screen.Height - FRMbulle.Height) / 2

    End If

    FRMbulle.Bulle.Style = LeStyle
    merlin.Balloon.Style = 0
    FRMbulleVisible = True
    FRMbulle.Bulle.CreateRegion
    FRMbulle.ReplaceControle
    FRMbulle.Show
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MNmerlinExit_Click()
    merlin.Stop
    If ChaineFin <> "" Then
        merlin.Speak ChaineFin
    End If
    Set fini = merlin.Hide
End Sub

'==================================================================================
'
'==================================================================================
Public Sub MNmerlinFreqLong_Click(index As Integer)
    Dim i As Long
    Dim mn
    For Each mn In MNmerlinFreqLong
        mn.Checked = False
    Next
    MNmerlinFreqLong(index).Checked = True
    Frequence = index
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MNmerlinHide_Click()


    If merlin.Visible Then
        lasttop = merlin.Top
        lastleft = merlin.Left
        merlin.Hide
    Else
        merlin.Top = lasttop
        merlin.Left = lastleft
        merlin.Show
        If ChaineBonjour <> "" Then
            merlin.Speak ChaineBonjour
        End If
    End If

End Sub

'==================================================================================
'
'==================================================================================
Private Sub MontreToi()
    If Not merlin.Visible Then merlin.Show
End Sub

'==================================================================================
'
'==================================================================================
Public Sub MNmerlinSizeBig_Click()
    merlin.Height = merlin.OriginalHeight * 2
    merlin.Width = merlin.OriginalWidth * 2
    MNmerlinSizeBig.Checked = True
    MNmerlinSizeNormal.Checked = False
    MNmerlinSizeSmall.Checked = False
End Sub

'==================================================================================
'
'==================================================================================
Public Sub MNmerlinSizeNormal_Click()
    merlin.Height = merlin.OriginalHeight
    merlin.Width = merlin.OriginalWidth
    MNmerlinSizeBig.Checked = False
    MNmerlinSizeNormal.Checked = True
    MNmerlinSizeSmall.Checked = False
End Sub

'==================================================================================
'
'==================================================================================
Public Sub MNmerlinSizeSmall_Click()
    merlin.Height = merlin.OriginalHeight / 2
    merlin.Width = merlin.OriginalWidth / 2
    MNmerlinSizeBig.Checked = False
    MNmerlinSizeNormal.Checked = False
    MNmerlinSizeSmall.Checked = True
End Sub

'==================================================================================
'
'==================================================================================
Public Sub MNmerlinTime_Click()
    Dim h As String
    Dim m As String
    Dim HideMe As Boolean
    Dim SayTime As String

    If Not merlin.Visible Then HideMe = True
    merlin.Top = lasttop
    merlin.Left = lastleft
    merlin.Show
    h = Mid(CStr(Time), 1, 2)
    m = Mid(CStr(Time), 4, 2)
    If Mid(h, 1, 1) = "0" Then h = Mid(h, 2, 1)
    If Mid(m, 1, 1) = "0" Then m = Mid(m, 2, 1)
    SayTime = Replace(ChaineTime, "#H", CStr(h))
    SayTime = Replace(SayTime, "#M", CStr(m))

    merlin.Speak SayTime
    If HideMe Then merlin.Hide
End Sub

'==================================================================================
'
'==================================================================================
Private Sub AfficheMessage()
On Error Resume Next
    Dim LeStyle  As Long

    LeStyle = StyleEnFonction(lastleft + merlin.Width / 2, lasttop + merlin.Height / 2)
    With FRMmessage
        If LeStyle = 0 Then
            .Top = lasttop * Screen.TwipsPerPixelY - .Height + 150
            .Left = lastleft * Screen.TwipsPerPixelX - .Width + merlin.Width * 5
        End If
        If LeStyle = 1 Then
            .Top = lasttop * Screen.TwipsPerPixelY - .Height + 150
            .Left = lastleft * Screen.TwipsPerPixelX + merlin.Width * 10
        End If
    
        If LeStyle = 4 Then
            .Top = lasttop * Screen.TwipsPerPixelY + merlin.Height * 15
            .Left = lastleft * Screen.TwipsPerPixelX + merlin.Width * 10
        End If
    
        If LeStyle = 5 Then
            .Top = lasttop * Screen.TwipsPerPixelY + merlin.Height * 15
            .Left = lastleft * Screen.TwipsPerPixelX - .Width + merlin.Width * 5
        End If
        .Bulle.Style = LeStyle
    End With
    
    'FRMmessage.Top = Screen.Height - merlin.Height * 15 - FRMmessage.Height
    'FRMmessage.Left = Screen.Width - merlin.Width * 8 - FRMmessage.Width
    
    FRMmessage.Show
    If Err.Number = 0 Then FRMmessage.Bulle.CreateRegion
End Sub

Private Sub MNmin_Click(index As Integer)
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Timer1_Timer()


    If Now >= NextMessageDate Then
        AfficheMessage
    End If


    If Frequence = 0 Then Exit Sub
    If m <> Mid(CStr(Time), 4, 2) Then
        m = Mid(CStr(Time), 4, 2)
        If m = "00" Then MNmerlinTime_Click
        If Frequence = 60 Then Exit Sub
        If m = "30" Then MNmerlinTime_Click
        If Frequence = 30 Then Exit Sub
        If m = "20" Or m = "40" Then MNmerlinTime_Click
        If Frequence = 20 Then Exit Sub
        If m = "15" Or m = "45" Then MNmerlinTime_Click
        If Frequence = 15 Then Exit Sub
        If m = "10" Or m = "40" Or m = "50" Then MNmerlinTime_Click
        If Frequence = 10 Then Exit Sub
        If m = "05" Or m = "25" Or m = "35" Or m = "55" Then MNmerlinTime_Click
    End If
End Sub

'==================================================================================
'
'==================================================================================
Public Sub GetOtherAgent()
Dim tempo As String
On Error GoTo aucun
    CommonDialog1.Filter = "Microsoft Agent Characters (*.acs)|*.acs"
    
    CommonDialog1.InitDir = fso.GetSpecialFolder(WindowsFolder) & "\msagent\chars"
    
    CommonDialog1.ShowOpen
    CommonDialog1.DialogTitle = "Choose an Agent ..."
    CommonDialog1.CancelError = True
    tempo = CommonDialog1.FileName
    If fso.FileExists(tempo) Then
        If InStr(1, UCase(tempo), ".ACS") <> 0 Then
            chaineagent = tempo
            SaveSetting "TimeAgent", "string", "Agent", chaineagent
            Agent1.Characters.Unload "merlin"
            Agent1.Characters.Load "merlin", chaineagent
            Set merlin = Agent1.Characters("merlin")
            merlin.SoundEffectsOn = True
            merlin.LanguageID = &H409
            merlin.AutoPopupMenu = False
            merlin.Show
        End If
    End If
aucun:
End Sub

