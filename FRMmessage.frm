VERSION 5.00
Object = "{4034C11D-8602-11D1-9840-002078110E7D}#1.0#0"; "ASASSISTANTPOPUP.OCX"
Object = "{D19CC187-8393-11D1-983F-002078110E7D}#1.0#0"; "ASBUBBLEFORM.OCX"
Begin VB.Form FRMmessage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin asBubbleWindow.asBubbleForm Bulle 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   3413
      BackColor       =   13434879
      Begin asAssistantPopup.asAssisPopup BTok 
         Height          =   435
         Left            =   4020
         TabIndex        =   9
         Top             =   1020
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         BackColor       =   13434879
         Caption         =   "Ok Close"
         Picture         =   "FRMmessage.frx":0000
         MouseOverPicture=   "FRMmessage.frx":059A
         MouseDownPicture=   "FRMmessage.frx":0B34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame FrameAbout 
         BackColor       =   &H00CCFFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   4995
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "fredjust@hotmail.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   660
            MouseIcon       =   "FRMmessage.frx":10CE
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   720
            Width           =   2625
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.fredjust.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   1080
            MouseIcon       =   "FRMmessage.frx":13D8
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   60
            Width           =   3270
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CCFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1215
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4995
         Begin VB.ComboBox CBremind 
            BackColor       =   &H00CCFFFF&
            Height          =   315
            ItemData        =   "FRMmessage.frx":16E2
            Left            =   1980
            List            =   "FRMmessage.frx":1710
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TXTmessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DFFFFF&
            Height          =   465
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "FRMmessage.frx":1794
            Top             =   180
            Width           =   4695
         End
         Begin asAssistantPopup.asAssisPopup BTremind 
            Height          =   435
            Left            =   360
            TabIndex        =   3
            Top             =   660
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   767
            BackColor       =   13434879
            Caption         =   "Remind me in :"
            Picture         =   "FRMmessage.frx":17AE
            MouseOverPicture=   "FRMmessage.frx":1D48
            MouseDownPicture=   "FRMmessage.frx":22E2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "01/01/2000"
            Height          =   195
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   870
         End
      End
   End
End
Attribute VB_Name = "FRMmessage"
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

Dim lastx As Long
Dim lasty As Long

'==================================================================================
'
'==================================================================================
Private Sub BTok_Click()
Dim ligne As ListItem

If ShowAbout Then
    merlin.Balloon.Style = 1
    Unload Me
    merlin.Hide
Else
On Error GoTo gestion_erreur

    Set ligne = FRMbulle.LV.ListItems(IndexLigne)
    
    Select Case ligne.SubItems(5)
        Case "One Time"
            FRMbulle.LV.ListItems.Remove IndexLigne
        Case "Daily"
            NextMessageDate = DateAdd("d", 1, Now)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
        Case "Weekly"
            NextMessageDate = DateAdd("d", 7, NextMessageDate)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
        Case "Monthly"
            NextMessageDate = DateAdd("m", 1, NextMessageDate)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
        Case "Yearly"
            NextMessageDate = DateAdd("yyyy", 1, NextMessageDate)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
    End Select
gestion_erreur:
    FRMbulle.ChercheMessageSuivant
    merlin.Balloon.Style = 1
    Unload Me
    merlin.Hide
End If
End Sub

'==================================================================================
' Créer le %date%
' Frédéric Just
'
'==================================================================================
Private Function ValeurINT(ByVal index As Long) As Long
    Select Case index
        Case 0
            ValeurINT = 5
        Case 1
            ValeurINT = 10
        Case 2
            ValeurINT = 15
        Case 3
            ValeurINT = 30
        Case 4
            ValeurINT = 60
        Case 5
            ValeurINT = 120
        Case 6
            ValeurINT = 180
        Case 7
            ValeurINT = 240
        Case 8
            ValeurINT = 480
        Case 9
            ValeurINT = 1
        Case 10
            ValeurINT = 2
        Case 11
            ValeurINT = 3
        Case 12
            ValeurINT = 4
        Case 13
            ValeurINT = 7
        Case Else
    End Select
End Function
'==================================================================================
'
'==================================================================================
Private Sub BTremind_Click()
    Dim ligne As ListItem
    'Dim TempoM As String
    
    Set ligne = FRMbulle.LV.ListItems(IndexLigne)
    NextMessageDate = DateAdd(IIf(CBremind.ListIndex > 8, "d", "n"), ValeurINT(CBremind.ListIndex), Now)
    'TempoM = TXTmessage
    
    Select Case ligne.SubItems(5)
        Case "One Time" ' update message
            ligne.Text = Day(NextMessageDate)
            
        Case Else ' add a new message
            Set ligne = FRMbulle.LV.ListItems.Add(, , Day(NextMessageDate))
            ligne.SubItems(5) = "One Time"
            ligne.Checked = True
            ligne.SubItems(7) = "1"
    End Select
    
    ligne.SubItems(1) = Month(NextMessageDate)
    ligne.SubItems(2) = Year(NextMessageDate)
    ligne.SubItems(3) = Hour(NextMessageDate)
    ligne.SubItems(4) = Minute(NextMessageDate)
    ligne.SubItems(6) = TXTmessage
    
    Set ligne = FRMbulle.LV.ListItems(IndexLigne)
    NextMessageDate = DateAdd("n", -5, NextMessageDate)
    
    Select Case ligne.SubItems(5)
        Case "One Time"
            'FRMbulle.LV.ListItems.Remove IndexLigne
        Case "Daily"
            NextMessageDate = DateAdd("d", 1, NextMessageDate)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
        Case "Weekly"
            NextMessageDate = DateAdd("d", 7, NextMessageDate)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
        Case "Monthly"
            NextMessageDate = DateAdd("m", 1, NextMessageDate)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
        Case "Yearly"
            NextMessageDate = DateAdd("yyyy", 1, NextMessageDate)
            ligne.Text = Day(NextMessageDate)
            ligne.SubItems(1) = Month(NextMessageDate)
            ligne.SubItems(2) = Year(NextMessageDate)
    End Select
    
    
    FRMbulle.ChercheMessageSuivant
    Unload Me
    merlin.Hide
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Activate()
    merlin.Balloon.Style = 0
    If Not merlin.Visible Then
        merlin.Show
        merlin.Balloon.Style = 0
    End If
    If Not ShowAbout Then merlin.Speak TXTmessage
    FRMmessageVisible = True
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Load()
    FRMmerlin.Timer1.Enabled = False
    merlin.Show
    
    If ShowAbout Then
        Frame1.Visible = False
        FrameAbout.Visible = True
        merlin.Speak "Time Agent" & Chr(13) & _
            "Version " & CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision) & Chr(13) & _
            "Fred Just Soft" & Chr(13) & Chr(13) & _
            "fredjust@hotmail point com"
    Else
        Frame1.Visible = True
        FrameAbout.Visible = False
        Label1 = NextMessageDate
        TXTmessage = FRMbulle.LV.ListItems(IndexLigne).SubItems(6)
        CBremind.ListIndex = 0
    End If
    
    merlin.Balloon.Style = 0
    FRMmessage.Width = Bulle.Width
    FRMmessage.Height = Bulle.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    merlin.Balloon.Style = 1
    FRMmerlin.Timer1.Enabled = True
    FRMmessageVisible = False
    ShowAbout = False
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lastx = X
    lasty = Y
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With FRMmessage
            .Left = .Left - (lastx - X)
            .Top = .Top - (lasty - Y)
        End With
    End If
End Sub

Private Sub Label2_Click()
    ShellEx "http://www.fredjust.com"
End Sub

Private Sub Label3_Click()
    ShellEx "mailto:fredjust@hotmail.com"
End Sub
