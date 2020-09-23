VERSION 5.00
Object = "{4034C11D-8602-11D1-9840-002078110E7D}#1.0#0"; "ASASSISTANTPOPUP.OCX"
Object = "{D19CC187-8393-11D1-983F-002078110E7D}#1.0#0"; "ASBUBBLEFORM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMbulle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9060
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMbulle.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMbulle.frx":015A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin asBubbleWindow.asBubbleForm Bulle 
      Height          =   4935
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   8705
      BackColor       =   13434879
      Style           =   5
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00CCFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4635
         Left            =   60
         TabIndex        =   13
         ToolTipText     =   "Move me, By Draq & Drop"
         Top             =   60
         Width           =   8535
         Begin asAssistantPopup.asAssisPopup BTabout 
            Height          =   435
            Left            =   6900
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   3420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   767
            BackColor       =   13434879
            Caption         =   "About"
            Picture         =   "FRMbulle.frx":02B4
            MouseOverPicture=   "FRMbulle.frx":084E
            MouseDownPicture=   "FRMbulle.frx":0DE8
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
         Begin VB.TextBox TXTmessage 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   300
            TabIndex        =   7
            Text            =   "Enter your message here"
            Top             =   4020
            Width           =   6735
         End
         Begin asAssistantPopup.asAssisPopup BTsay 
            Height          =   375
            Left            =   7140
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BackColor       =   13434879
            Caption         =   "SAY"
            Picture         =   "FRMbulle.frx":1382
            MouseOverPicture=   "FRMbulle.frx":191C
            MouseDownPicture=   "FRMbulle.frx":1EB6
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
         Begin asAssistantPopup.asAssisPopup BTok 
            Height          =   375
            Left            =   7260
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   105
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   661
            BackColor       =   13434879
            Caption         =   "OK"
            Picture         =   "FRMbulle.frx":2450
            MouseOverPicture=   "FRMbulle.frx":29EA
            MouseDownPicture=   "FRMbulle.frx":2F84
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
         Begin asAssistantPopup.asAssisPopup BTtime 
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   105
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BackColor       =   13434879
            Caption         =   "Option"
            Picture         =   "FRMbulle.frx":351E
            MouseOverPicture=   "FRMbulle.frx":3AB8
            MouseDownPicture=   "FRMbulle.frx":4052
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
         Begin asAssistantPopup.asAssisPopup BTrdv 
            Height          =   375
            Left            =   240
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   105
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BackColor       =   13434879
            Caption         =   "RDV"
            Picture         =   "FRMbulle.frx":45EC
            MouseOverPicture=   "FRMbulle.frx":4B86
            MouseDownPicture=   "FRMbulle.frx":5120
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
         Begin VB.Frame FrameRDV 
            Appearance      =   0  'Flat
            BackColor       =   &H00CCFFFF&
            Caption         =   "RDV"
            ForeColor       =   &H80000008&
            Height          =   3915
            Left            =   180
            TabIndex        =   36
            Top             =   540
            Width           =   8175
            Begin VB.ComboBox CBfreq 
               Height          =   315
               ItemData        =   "FRMbulle.frx":56BA
               Left            =   1560
               List            =   "FRMbulle.frx":56CD
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   3060
               Width           =   2175
            End
            Begin VB.TextBox TXTyear 
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   780
               TabIndex        =   3
               Text            =   "2000"
               Top             =   3120
               Width           =   555
            End
            Begin VB.TextBox TXTmois 
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   780
               TabIndex        =   2
               Text            =   "01"
               Top             =   2760
               Width           =   555
            End
            Begin VB.TextBox TXTday 
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   780
               TabIndex        =   1
               Text            =   "01"
               Top             =   2400
               Width           =   555
            End
            Begin VB.TextBox TXTheure 
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   2280
               TabIndex        =   4
               Text            =   "12"
               Top             =   2400
               Width           =   555
            End
            Begin VB.TextBox TXTminute 
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   2280
               TabIndex        =   5
               Text            =   "00"
               Top             =   2760
               Width           =   555
            End
            Begin asAssistantPopup.asAssisPopup BTdelete 
               Height          =   435
               Left            =   6720
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   2340
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   767
               BackColor       =   13434879
               Caption         =   "Delete"
               Picture         =   "FRMbulle.frx":56FB
               MouseOverPicture=   "FRMbulle.frx":5C95
               MouseDownPicture=   "FRMbulle.frx":622F
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
            Begin MSComctlLib.ListView LV 
               Height          =   2055
               Left            =   120
               TabIndex        =   0
               Tag             =   "0"
               Top             =   240
               Width           =   7875
               _ExtentX        =   13891
               _ExtentY        =   3625
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               SmallIcons      =   "ImageList1"
               ForeColor       =   -2147483640
               BackColor       =   14680063
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Day"
                  Object.Width           =   1350
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Month"
                  Object.Width           =   1191
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Year"
                  Object.Width           =   1005
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Hr:"
                  Object.Width           =   970
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Min."
                  Object.Width           =   970
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Freq."
                  Object.Width           =   1614
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Message"
                  Object.Width           =   6165
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Active"
                  Object.Width           =   0
               EndProperty
            End
            Begin asAssistantPopup.asAssisPopup BTadd 
               Height          =   435
               Left            =   4320
               TabIndex        =   8
               Top             =   2340
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   767
               BackColor       =   13434879
               Caption         =   "Add"
               Picture         =   "FRMbulle.frx":67C9
               MouseOverPicture=   "FRMbulle.frx":6D63
               MouseDownPicture=   "FRMbulle.frx":72FD
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
            Begin asAssistantPopup.asAssisPopup asAssisPopup1 
               Height          =   435
               Left            =   4320
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   2880
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   767
               BackColor       =   13434879
               Caption         =   "UpDate"
               Picture         =   "FRMbulle.frx":7897
               MouseOverPicture=   "FRMbulle.frx":7E31
               MouseDownPicture=   "FRMbulle.frx":83CB
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
            Begin asAssistantPopup.asAssisPopup BTload 
               Height          =   435
               Left            =   5520
               TabIndex        =   45
               Top             =   2340
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   767
               BackColor       =   13434879
               Caption         =   "Load"
               Picture         =   "FRMbulle.frx":8965
               MouseOverPicture=   "FRMbulle.frx":8EFF
               MouseDownPicture=   "FRMbulle.frx":9499
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
            Begin asAssistantPopup.asAssisPopup asAssisPopup2 
               Height          =   435
               Left            =   5520
               TabIndex        =   46
               Top             =   2880
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   767
               BackColor       =   13434879
               Caption         =   "Save as"
               Picture         =   "FRMbulle.frx":9A33
               MouseOverPicture=   "FRMbulle.frx":9FCD
               MouseDownPicture=   "FRMbulle.frx":A567
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Year :"
               Height          =   195
               Left            =   180
               TabIndex        =   42
               Top             =   3120
               Width           =   420
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Month :"
               Height          =   195
               Left            =   180
               TabIndex        =   41
               Top             =   2760
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Day :"
               Height          =   195
               Left            =   180
               TabIndex        =   40
               Top             =   2400
               Width           =   375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hour :"
               Height          =   195
               Left            =   1560
               TabIndex        =   39
               Top             =   2400
               Width           =   435
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Minute :"
               Height          =   195
               Left            =   1560
               TabIndex        =   38
               Top             =   2760
               Width           =   570
            End
         End
         Begin VB.Frame FrameOther 
            BackColor       =   &H00CCFFFF&
            Caption         =   "Option"
            Height          =   3915
            Left            =   180
            TabIndex        =   18
            Top             =   540
            Visible         =   0   'False
            Width           =   8175
            Begin VB.Frame Frame1 
               BackColor       =   &H00CCFFFF&
               Caption         =   "Say Time ..."
               Height          =   3135
               Left            =   180
               TabIndex        =   27
               Top             =   240
               Width           =   3135
               Begin VB.OptionButton OptionFreq 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Each Hour"
                  Height          =   315
                  Index           =   60
                  Left            =   120
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   1680
                  Width           =   1395
               End
               Begin VB.OptionButton OptionFreq 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Each 30 minutes"
                  Height          =   315
                  Index           =   30
                  Left            =   120
                  TabIndex        =   32
                  TabStop         =   0   'False
                  Top             =   1320
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton OptionFreq 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Each 15 minutes"
                  Height          =   315
                  Index           =   15
                  Left            =   120
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   960
                  Width           =   1575
               End
               Begin VB.OptionButton OptionFreq 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Each 5 minutes"
                  Height          =   315
                  Index           =   5
                  Left            =   120
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.OptionButton OptionFreq 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Never !"
                  Height          =   315
                  Index           =   0
                  Left            =   120
                  TabIndex        =   29
                  TabStop         =   0   'False
                  Top             =   2040
                  Width           =   1575
               End
               Begin VB.TextBox TXTtime 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   120
                  TabIndex        =   10
                  Text            =   "It is #H hours, #M minutes."
                  Top             =   2700
                  Width           =   2895
               End
               Begin VB.OptionButton OptionFreq 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Each 10 minutes"
                  Height          =   315
                  Index           =   10
                  Left            =   120
                  TabIndex        =   28
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   1575
               End
               Begin asAssistantPopup.asAssisPopup BTtest 
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   2220
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  BackColor       =   13434879
                  Caption         =   "Test"
                  Picture         =   "FRMbulle.frx":AB01
                  MouseOverPicture=   "FRMbulle.frx":B09B
                  MouseDownPicture=   "FRMbulle.frx":B635
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
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Format :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   35
                  Top             =   2460
                  Width           =   570
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00CCFFFF&
               Caption         =   "First Message "
               Height          =   615
               Left            =   3600
               TabIndex        =   25
               Top             =   240
               Width           =   4275
               Begin VB.TextBox TXTbonjour 
                  Appearance      =   0  'Flat
                  Height          =   255
                  Left            =   120
                  TabIndex        =   11
                  Text            =   "Yes Master !"
                  Top             =   240
                  Width           =   2895
               End
               Begin asAssistantPopup.asAssisPopup BTfirst 
                  Height          =   375
                  Left            =   3120
                  TabIndex        =   26
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  BackColor       =   13434879
                  Caption         =   "Test"
                  Picture         =   "FRMbulle.frx":BBCF
                  MouseOverPicture=   "FRMbulle.frx":C169
                  MouseDownPicture=   "FRMbulle.frx":C703
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
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00CCFFFF&
               Caption         =   "End Message "
               Height          =   615
               Left            =   3600
               TabIndex        =   23
               Top             =   960
               Width           =   4275
               Begin VB.TextBox TXTfin 
                  Appearance      =   0  'Flat
                  Height          =   255
                  Left            =   120
                  TabIndex        =   12
                  Text            =   "Your desire is an order."
                  Top             =   240
                  Width           =   2895
               End
               Begin asAssistantPopup.asAssisPopup BTfin 
                  Height          =   375
                  Left            =   3120
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  BackColor       =   13434879
                  Caption         =   "Test"
                  Picture         =   "FRMbulle.frx":CC9D
                  MouseOverPicture=   "FRMbulle.frx":D237
                  MouseDownPicture=   "FRMbulle.frx":D7D1
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
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00CCFFFF&
               Caption         =   "Agent Size"
               Height          =   1575
               Left            =   3600
               TabIndex        =   19
               Top             =   1740
               Width           =   1275
               Begin VB.OptionButton OPsize 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Small"
                  Height          =   375
                  Index           =   0
                  Left            =   180
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton OPsize 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Normal"
                  Height          =   375
                  Index           =   1
                  Left            =   180
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   660
                  Value           =   -1  'True
                  Width           =   1035
               End
               Begin VB.OptionButton OPsize 
                  BackColor       =   &H00CCFFFF&
                  Caption         =   "Big"
                  Height          =   375
                  Index           =   2
                  Left            =   180
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   975
               End
            End
            Begin asAssistantPopup.asAssisPopup BTchange 
               Height          =   375
               Left            =   6120
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   1800
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   661
               BackColor       =   13434879
               Caption         =   "Change Agent ..."
               Picture         =   "FRMbulle.frx":DD6B
               MouseOverPicture=   "FRMbulle.frx":E305
               MouseDownPicture=   "FRMbulle.frx":E89F
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
         End
         Begin VB.Shape Shape1 
            Height          =   465
            Left            =   180
            Shape           =   4  'Rounded Rectangle
            Top             =   60
            Width           =   8175
         End
      End
   End
End
Attribute VB_Name = "FRMbulle"
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

Dim ClickBT As Boolean

'==================================================================================
'
'==================================================================================
Private Sub asAssisPopup1_Click()
    Dim ligne As ListItem
    Dim tempo As String
    Dim tempodate As Date
    
    tempo = TXTday & "/" & TXTmois & "/" & TXTyear & " " & TXTheure & ":" & TXTminute
    On Error Resume Next
        tempodate = Format(tempo, "DD/MM/YYYY HH:MM")
    If Err <> 0 Then
        merlin.Speak "Sorry but, it's a wrong date."
        Exit Sub
    Else
        merlin.Speak "Ok."
    End If

    Set ligne = LV.SelectedItem
    If Not ligne Is Nothing Then
        ligne.Text = TXTday
        ligne.SubItems(1) = TXTmois
        ligne.SubItems(2) = TXTyear
        ligne.SubItems(3) = TXTheure
        ligne.SubItems(4) = TXTminute
        ligne.Tag = CBfreq.ListIndex
        ligne.SubItems(5) = CBfreq
        ligne.SubItems(6) = TXTmessage
        ligne.SubItems(7) = "1"
        ligne.Checked = True
            
        ChercheMessageSuivant
    End If
End Sub

Private Sub asAssisPopup2_Click()
    ClickBT = True

End Sub

Private Sub asAssisPopup2_MouseExit()
    Dim tempo As String
    If ClickBT Then
        tempo = EnregistrerSous("LV")
        If tempo <> "" Then
            SaveAs = tempo
            FrameRDV.Caption = "RDV : " & SaveAs
            EcritContenuListViewDansFichier LV, SaveAs
        End If
        ClickBT = False
    End If
End Sub

Private Sub BTabout_Click()
    ClickBT = True
    
End Sub

Private Sub BTabout_MouseExit()
    If ClickBT Then
        If Not merlin.Visible Then merlin.Show
        merlin.Speak "Time Agent" & Chr(13) & _
            "Version " & CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision) & Chr(13) & _
            "Fred Just Soft" & Chr(13) & Chr(13) & _
            "fredjust@hotmail point com"
        
        MsgBox "Time Agent" & Chr(13) & _
            "Version " & CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision) & Chr(13) & _
            "Fred Just Soft" & Chr(13) & Chr(13) & _
            "fredjust@hotmail.com", vbInformation, "About Time Agent"
        ClickBT = False
    End If
        
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTadd_Click()
    Dim ligne As ListItem
    Dim tempo As String
    Dim tempodate As Date
    
    tempo = TXTday & "/" & TXTmois & "/" & TXTyear & " " & TXTheure & ":" & TXTminute
    On Error Resume Next
        tempodate = Format(tempo, "DD/MM/YYYY HH:MM")
    If Err <> 0 Then
        merlin.Speak "Sorry but, it's a wrong date."
        Exit Sub
    Else
        merlin.Speak "Ok."
    End If
    
    Set ligne = LV.ListItems.Add(, , TXTday)
    ligne.SubItems(1) = TXTmois
    ligne.SubItems(2) = TXTyear
    ligne.SubItems(3) = TXTheure
    ligne.SubItems(4) = TXTminute
    ligne.Tag = CBfreq.ListIndex
    ligne.SubItems(5) = CBfreq
    ligne.SubItems(6) = TXTmessage
    ligne.SubItems(7) = "1"
    ligne.Checked = True
    
    
    
    ChercheMessageSuivant
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTchange_Click()
ClickBT = True
End Sub

Private Sub BTchange_MouseExit()
    If ClickBT Then
        FRMmerlin.GetOtherAgent
        ClickBT = False
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTdelete_Click()
    If Not LV.SelectedItem Is Nothing Then LV.ListItems.Remove LV.SelectedItem.index
    ChercheMessageSuivant
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTfin_Click()
    If ChaineFin <> "" Then
        If Not merlin.Visible Then merlin.Show
        merlin.Speak ChaineFin
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTfirst_Click()
    If ChaineBonjour <> "" Then
        If Not merlin.Visible Then merlin.Show
        merlin.Speak ChaineBonjour
    End If
End Sub

Private Sub BTload_Click()
    ClickBT = True
End Sub

Private Sub BTload_MouseExit()
Dim tempo As String
    If ClickBT Then
        tempo = OuvrirFichierExistant("LV", , "Choose a File")
            If tempo <> "" Then
                SaveAs = tempo
                FrameRDV.Caption = "RDV : " & SaveAs
                LisContenuLISTVIEWdepuisFichier LV, SaveAs
            End If
        ClickBT = False
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTok_Click()
    ClickBT = True
    
End Sub

'==================================================================================
'
'==================================================================================
Public Sub ChercheMessageSuivant()
    Dim ligne As ListItem
    Dim tempodate As Date
    Dim tempo As String
    Dim NextDate As Date
    
    On Error Resume Next
    
    NextDate = Format("12/12/2099", "DD/MM/YYYY")
    
    For Each ligne In LV.ListItems
        tempo = ligne.Text & "/" & ligne.SubItems(1) & "/" & ligne.SubItems(2) & " " & ligne.SubItems(3) & ":" & ligne.SubItems(4)
        tempodate = Format(tempo, "DD/MM/YYYY HH:MM")
        ligne.SmallIcon = 1
        If NextDate > tempodate And ligne.Checked Then
            NextDate = tempodate
            IndexLigne = ligne.index
            ligne.Selected = True
            ligne.EnsureVisible
        End If
    Next
    
    LV.ListItems(IndexLigne).SmallIcon = 2
    NextMessageDate = NextDate
    EcritContenuListViewDansFichier LV, SaveAs
End Sub

Private Sub BTok_MouseExit()
    If ClickBT Then
        Me.Hide
        merlin.Balloon.Style = 1
        merlin.Hide
        FRMbulleVisible = False
        ChercheMessageSuivant
        ClickBT = False
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTrdv_Click()
    FrameOther.Visible = False
    FrameRDV.Visible = True
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTsay_Click()
    If TXTmessage <> "" Then
        If Not merlin.Visible Then merlin.Show
        merlin.Speak TXTmessage
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTtest_Click()
    FRMmerlin.MNmerlinTime_Click
End Sub

'==================================================================================
'
'==================================================================================
Private Sub BTtime_Click()
    FrameOther.Visible = True
    FrameRDV.Visible = False
End Sub

'==================================================================================
'
'==================================================================================
Public Sub ReplaceControle()
    Select Case Bulle.Style
        Case 0, 1
            Frame.Top = 60
            Frame.Left = 60
            Bulle.Width = Frame.Width + 120
            Bulle.Height = Frame.Height + 120 + 300
        Case 2, 3
            Frame.Top = 60
            Frame.Left = 360
            Bulle.Width = Frame.Width + 120 + 300
            Bulle.Height = Frame.Height + 120
        Case 4, 5
            Frame.Top = 360
            Frame.Left = 60
            Bulle.Width = Frame.Width + 120
            Bulle.Height = Frame.Height + 120 + 300
        Case 6, 7
            Frame.Top = 60
            Frame.Left = 60
            Bulle.Width = Frame.Width + 120 + 300
            Bulle.Height = Frame.Height + 120
    End Select
    FRMbulle.Width = Bulle.Width
    FRMbulle.Height = Bulle.Height

End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        merlin.Balloon.Style = 1
        Me.Hide
        FRMbulleVisible = False
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Load()
    InMem = True

    TXTday = Day(Now)
    TXTmois = Month(Now)
    TXTyear = Year(Now)
    TXTminute = Minute(Now)
    TXTheure = Hour(Now)
    
    
    LisContenuLISTVIEWdepuisFichier FRMbulle.LV, SaveAs
    ChercheMessageSuivant
    Bulle.Style = 1
    ReplaceControle
    FrameRDV.Caption = "RDV : " & SaveAs
    Bulle.CreateRegion


    CBfreq.ListIndex = 1
    TXTbonjour = ChaineBonjour
    TXTfin = ChaineFin
    TXTtime = ChaineTime
    OptionFreq(Frequence).Value = True
    OPsize(Round(merlin.Height / merlin.OriginalHeight - 0.1)).Value = True


End Sub

'==================================================================================
'
'==================================================================================
Private Sub Frame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lastx = X
    lasty = Y
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Frame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        FRMbulle.Left = FRMbulle.Left - (lastx - X)
        FRMbulle.Top = FRMbulle.Top - (lasty - Y)
    End If
End Sub


'==================================================================================
'
'==================================================================================
Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ClasseLesColonnes LV, ColumnHeader
End Sub

'==================================================================================
'
'==================================================================================
Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.SubItems(7) = IIf(Item.Checked, "1", "0")
End Sub

'==================================================================================
'
'==================================================================================
Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TXTday = Item.Text
    TXTmois = Item.SubItems(1)
    TXTyear = Item.SubItems(2)
    TXTheure = Item.SubItems(3)
    TXTminute = Item.SubItems(4)
    CBfreq = Item.SubItems(5)
    TXTmessage = Item.SubItems(6)
End Sub

'==================================================================================
'
'==================================================================================
Private Sub OPsize_Click(index As Integer)
    Select Case index
        Case 0
            FRMmerlin.MNmerlinSizeSmall_Click
        Case 1
            FRMmerlin.MNmerlinSizeNormal_Click
        Case 2
            FRMmerlin.MNmerlinSizeBig_Click
    End Select
End Sub

'==================================================================================
'
'==================================================================================
Private Sub OptionFreq_Click(index As Integer)
    FRMmerlin.MNmerlinFreqLong_Click index
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTbonjour_LostFocus()
    ChaineBonjour = Trim(TXTbonjour)
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTday_GotFocus()
    SelectAll TXTday
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTfin_LostFocus()
    ChaineFin = Trim(TXTfin)
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTheure_GotFocus()
    SelectAll TXTheure
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTmessage_GotFocus()
    SelectAll TXTmessage
End Sub

Private Sub TXTmessage_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then BTsay_Click
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTmessage_LostFocus()
    TXTmessage = Trim(TXTmessage)
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTminute_GotFocus()
    SelectAll TXTminute
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTmois_GotFocus()
    SelectAll TXTmois
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTtime_LostFocus()
    If InStr(1, TXTtime, "#H") = 0 And InStr(1, TXTtime, "#M") = 0 Then
        MsgBox "Invalide 'time format' !" + Chr(13) + _
                "Use #H for Hours and #M for Minutes" _
                + Chr(13) + "Samples : " & ChaineTime, vbInformation, "It's a Time Agent Error !"
        TXTtime.SetFocus
    Else
        ChaineTime = TXTtime
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub TXTyear_GotFocus()
    SelectAll TXTyear
End Sub
