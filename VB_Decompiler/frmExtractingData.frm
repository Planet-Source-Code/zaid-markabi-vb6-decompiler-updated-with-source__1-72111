VERSION 5.00
Begin VB.Form frmExtractingData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extracting Data"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   Icon            =   "frmExtractingData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3585
      ScaleWidth      =   9705
      TabIndex        =   4
      Top             =   240
      Width           =   9735
      Begin VB.Timer RefList 
         Enabled         =   0   'False
         Interval        =   2500
         Left            =   3600
         Top             =   1680
      End
      Begin VB.Timer TimerForExtract 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4320
         Top             =   1680
      End
      Begin VB.PictureBox LastImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   240
         Picture         =   "frmExtractingData.frx":08CA
         ScaleHeight     =   1065
         ScaleWidth      =   4545
         TabIndex        =   17
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox LastSentence 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2160
         Width           =   4575
      End
      Begin VB.PictureBox LoadingExtAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   345
         ScaleWidth      =   9225
         TabIndex        =   14
         Top             =   1320
         Width           =   9255
         Begin VB.PictureBox LoadingExtDone 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   15
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   7800
         Picture         =   "frmExtractingData.frx":09CC
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   12
         Top             =   680
         Width           =   1215
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STOP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7800
         Picture         =   "frmExtractingData.frx":1DA6
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   10
         Top             =   310
         Width           =   1215
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.TextBox TextMaxSentence 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3720
         TabIndex        =   8
         Text            =   "4"
         Top             =   640
         Width           =   615
      End
      Begin VB.CheckBox ChkSentences 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Extract Sentences ( Text )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox ChkImages 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Extract Images ( Bmp , Jpg )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   1920
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Sentence"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   5040
         TabIndex        =   18
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Shape Shape3 
         Height          =   855
         Left            =   7200
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*selected length will be accepted"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4440
         TabIndex        =   9
         Top             =   705
         Width           =   2340
      End
      Begin VB.Shape Shape2 
         Height          =   855
         Left            =   3480
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extracted Sentences Length  (Minium)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   240
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.FileListBox ExtractedFilesSentences 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3345
      Left            =   5160
      Pattern         =   "*.Txt*"
      TabIndex        =   2
      Top             =   4200
      Width           =   4815
   End
   Begin VB.FileListBox ExtractedFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3345
      Left            =   240
      Pattern         =   "*.bmp*;*.jpg*"
      TabIndex        =   0
      Top             =   4200
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extracted Sentences"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   3
      Left            =   5280
      TabIndex        =   3
      Top             =   3960
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extracted Images"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Width           =   1485
   End
End
Attribute VB_Name = "frmExtractingData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim I As Integer
ExtractedFiles.Path = App.Path + "\Temp\"
ExtractedFilesSentences.Path = App.Path + "\Temp\"
For I = 0 To ExtractedFiles.ListCount - 1
Kill App.Path + "\Temp\" + ExtractedFiles.List(I)
Next
For I = 0 To ExtractedFilesSentences.ListCount - 1
Kill App.Path + "\Temp\" + ExtractedFilesSentences.List(I)
Next
ExtractedFiles.Refresh
ExtractedFilesSentences.Refresh
LoadingExtDone.Width = 0
End Sub

Private Sub Label4_Click()
Picture2.Enabled = False
Picture3.Enabled = True
Open FileName For Binary As #1
FilePos = 1
MaxSentenceLong = Int(TextMaxSentence.Text) - 1
TimerForExtract.Enabled = True
RefList.Enabled = True
End Sub

Private Sub Label5_Click()
Close #1
End
End Sub

Private Sub RefList_Timer()
ExtractedFilesSentences.Refresh
ExtractedFiles.Refresh
End Sub

Private Sub TextMaxSentence_Change()
On Error GoTo Err:
If Int(TextMaxSentence.Text) < 999 And Int(TextMaxSentence.Text) > 0 Then
Exit Sub
End If
Err:
TextMaxSentence.Text = "4"
TextMaxSentence.SelStart = Len(TextMaxSentence.Text)
End Sub

Private Sub TimerForExtract_Timer()
Dim I As Integer
For I = 0 To 5000
Call ExtractNext
Next
LastSentence.Text = LastSavedSentence
RefreshLoadingBar
End Sub
