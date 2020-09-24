VERSION 5.00
Begin VB.Form frmEditor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10215
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2145
      ScaleWidth      =   4785
      TabIndex        =   34
      Top             =   6840
      Width           =   4815
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   3360
         TabIndex        =   39
         Top             =   150
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3360
         Picture         =   "frmEditor.frx":08CA
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   36
         Top             =   480
         Width           =   1215
         Begin VB.Label Label7 
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
            TabIndex        =   37
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WebSite : www.YazanMarkabi.Jeeran.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   44
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em@l : ZaidMarkabi@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   43
         Top             =   1560
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB Decompiler Lite ,      by   Zaid Markabi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   42
         Top             =   1320
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
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
         Index           =   13
         Left            =   240
         TabIndex        =   41
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Working ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4680
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decompile to new EXE file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   12
         Left            =   360
         TabIndex        =   38
         Top             =   480
         Width           =   2730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decompile"
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
         Index           =   11
         Left            =   240
         TabIndex        =   35
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox ScrSentences 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   5160
      Picture         =   "frmEditor.frx":1CA4
      ScaleHeight     =   5145
      ScaleWidth      =   4785
      TabIndex        =   5
      Top             =   3840
      Width           =   4815
      Begin VB.TextBox FindText 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   25
         Top             =   1920
         Width           =   2895
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3120
         Picture         =   "frmEditor.frx":1CEA
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Find"
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
            TabIndex        =   24
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.PictureBox LoadingExtAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   240
         ScaleHeight     =   105
         ScaleWidth      =   2625
         TabIndex        =   21
         Top             =   4680
         Width           =   2655
         Begin VB.PictureBox LoadingExtDone 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   0
            ScaleHeight     =   135
            ScaleWidth      =   375
            TabIndex        =   22
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3120
         Picture         =   "frmEditor.frx":30C4
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   19
         Top             =   4560
         Width           =   1215
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Replace"
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
            TabIndex        =   20
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.TextBox NewText 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   18
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox OrgText 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   17
         Top             =   3360
         Width           =   2895
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3120
         Picture         =   "frmEditor.frx":449E
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
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
      Begin VB.TextBox TxtSentences 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   8
         Top             =   360
         Width           =   4335
      End
      Begin VB.VScrollBar VScrSentences 
         Height          =   5175
         LargeChange     =   10
         Left            =   4560
         Max             =   100
         SmallChange     =   5
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Text for all sentences "
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
         Index           =   8
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   26
         Top             =   1920
         Width           =   810
      End
      Begin VB.Line Line2 
         X1              =   4440
         X2              =   120
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   16
         Top             =   3960
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orginal Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   3360
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace Text for all sentences "
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
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   2685
      End
      Begin VB.Line Line1 
         X1              =   4440
         X2              =   120
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sentences Editor"
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
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1470
      End
   End
   Begin VB.PictureBox ScrImages 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   240
      ScaleHeight     =   4.321
      ScaleMode       =   0  'User
      ScaleWidth      =   319
      TabIndex        =   4
      Top             =   3840
      Width           =   4815
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3000
         Picture         =   "frmEditor.frx":5878
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   30
         Top             =   2400
         Width           =   1215
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Import"
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
            TabIndex        =   31
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3000
         Picture         =   "frmEditor.frx":6C52
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export"
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
            TabIndex        =   29
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.PictureBox PicImages 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1425
         ScaleWidth      =   4305
         TabIndex        =   11
         Top             =   360
         Width           =   4335
      End
      Begin VB.VScrollBar VScrImages 
         Height          =   2895
         LargeChange     =   10
         Left            =   4560
         Max             =   100
         SmallChange     =   5
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace This Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   10
         Left            =   480
         TabIndex        =   33
         Top             =   2400
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export This Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   9
         Left            =   480
         TabIndex        =   32
         Top             =   1920
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Images Editor"
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
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1170
      End
   End
   Begin VB.FileListBox ExtractedFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3345
      Left            =   240
      Pattern         =   "*.bmp*;*.jpg*"
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin VB.FileListBox ExtractedFilesSentences 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3345
      Left            =   5160
      Pattern         =   "*.Txt*"
      TabIndex        =   0
      Top             =   360
      Width           =   4815
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
      TabIndex        =   3
      Top             =   120
      Width           =   1485
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
      TabIndex        =   2
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ExtractedFiles_Click()
VScrImages.Value = ExtractedFiles.ListIndex
VScrImages_Scroll
End Sub

Private Sub ExtractedFilesSentences_Click()
VScrSentences.Value = ExtractedFilesSentences.ListIndex
VScrSentences_Scroll
End Sub

Private Sub Form_Load()
ExtractedFiles.Path = App.Path + "\Temp\"
ExtractedFilesSentences.Path = App.Path + "\Temp\"
VScrImages.Max = ExtractedFiles.ListCount - 1
VScrSentences.Max = ExtractedFilesSentences.ListCount - 1
VScrImages_Scroll
VScrSentences_Scroll
LoadingExtDone.Width = 0
DoEvents
Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label2_Click()
On Error Resume Next
Dim I As Integer
Dim n As Integer
Dim X As String
NewText.MaxLength = Len(OrgText.Text)
NewText.Text = NewText.Text + Space(NewText.MaxLength - Len(NewText.Text))
If Not UCase(OrgText.Text) = UCase(NewText.Text) And Not OrgText.Text = "" Then
If MsgBox("Are you sure ?" + vbCrLf + vbCrLf + "* if you replaced text for all sentences, you may damage the program !" + vbCrLf + vbCrLf, vbYesNo + vbQuestion, "Replace") = vbYes Then
LoadingExtAll.Visible = True
For I = 0 To ExtractedFilesSentences.ListCount - 1
' open old text
Open App.Path + "\Temp\" + ExtractedFilesSentences.List(I) For Input As #1
Input #1, X
Close #1
' replace
Do While InStr(1, UCase(X), UCase(OrgText.Text)) > 0
n = InStr(1, UCase(X), UCase(OrgText.Text))
X = Left(X, n - 1) + NewText.Text + Right(X, Len(X) - n + 1 - Len(OrgText.Text))
Loop
' save new text
Open App.Path + "\Temp\" + ExtractedFilesSentences.List(I) For Binary As #2
Put #2, 1, X
Close #2
RefreshLoadingBarB I
DoEvents
Next
End If
End If
MsgBox vbCrLf + vbCrLf + Space(20) + "Done !" + Space(30) + vbCrLf + vbCrLf, vbInformation, "Finished !"
End Sub

Private Sub Label3_Click()
'On Error Resume Next
Dim I As Integer
Dim X As String
For I = ExtractedFilesSentences.ListIndex + 1 To ExtractedFilesSentences.ListCount - 1
Open App.Path + "\Temp\" + ExtractedFilesSentences.List(I) For Input As #1
Input #1, X
Close #1
If InStr(1, UCase(X), UCase(FindText.Text)) > 0 Then
ExtractedFilesSentences.ListIndex = I
VScrSentences.Value = I
Exit Sub
End If
Next
MsgBox vbCrLf + vbCrLf + Space(20) + "End !" + Space(30) + vbCrLf + vbCrLf, vbInformation, "Finished !"
End Sub

Private Sub Label4_Click()
TxtSentences.Text = TxtSentences.Text + Space(TxtSentences.MaxLength - Len(TxtSentences.Text))
Open App.Path + "\Temp\" + ExtractedFilesSentences.List(VScrSentences.Value) For Binary As #1
Put #1, 1, TxtSentences.Text
Close #1
End Sub

Private Sub Label5_Click()
On Error GoTo Err
SavePicture PicImages.Picture, "C:\Exported." + Right(ExtractedFiles.List(ExtractedFiles.ListIndex), 3)
MsgBox vbCrLf + vbCrLf + Space(20) + "Exported to " + "C:\Exported." + Right(ExtractedFiles.List(ExtractedFiles.ListIndex), 3) + Space(30) + vbCrLf + vbCrLf, vbInformation, "Exported !"
Err:
End Sub

Private Sub Label6_Click()
On Error GoTo Err
Dim Ln1 As Long
Dim Ln2 As Long
If MsgBox("Are you sure ?" + vbCrLf + vbCrLf + "* New image will be loaded from C:\Exported." + Right(ExtractedFiles.List(ExtractedFiles.ListIndex), 3) + vbCrLf + vbCrLf, vbYesNo + vbQuestion, "Replace") = vbYes Then
Open "C:\Exported." + Right(ExtractedFiles.List(ExtractedFiles.ListIndex), 3) For Input As #1
Ln1 = LOF(1)
Close #1
Open App.Path + "\Temp\" + ExtractedFiles.List(ExtractedFiles.ListIndex) For Input As #2
Ln2 = LOF(2)
Close #2
If Ln1 = Ln2 Then
FileCopy "C:\Exported." + Right(ExtractedFiles.List(ExtractedFiles.ListIndex), 3), App.Path + "\Temp\" + ExtractedFiles.List(ExtractedFiles.ListIndex)
PicImages.Picture = LoadPicture(App.Path + "\Temp\" + ExtractedFiles.List(ExtractedFiles.ListIndex))
MsgBox vbCrLf + vbCrLf + Space(20) + "Done !" + Space(30) + vbCrLf + vbCrLf, vbInformation, "Import !"
Else
MsgBox vbCrLf + vbCrLf + "Can't replace !" + Space(90) + vbCrLf + vbCrLf + " * New image file size should be smaller or equal the old file in the size" + vbCrLf + vbCrLf + Format(1 + Ln1 \ 1012, "0.0") + " Kb =/= " + Format(1 + Ln2 \ 1012, "0.0") + " Kb", vbInformation, "Import !"
End If
End If
Err:
End Sub

Private Sub Label7_Click()
On Error GoTo Err
Label8.Visible = True
DoEvents
Dim FileTitle As String
Dim I As Integer
File1.Path = App.Path + "\Temp\"
File1.Refresh
Open App.Path + "\OrginalFile.dat" For Input As #1
Input #1, FileTitle
Close #1
FileCopy App.Path + "\" + FileTitle + ".exe", "C:\Decompiled.exe"
Open "C:\Decompiled.exe" For Binary As #1
For I = 0 To File1.ListCount - 1
Open File1.Path + "\" + File1.List(I) For Binary As #2
DoEvents
If LOF(2) > 0 Then
ReDim CHARn(LOF(2) - 1)
Get #2, 1, CHARn
Put #1, CLng(Mid(File1.List(I), Len(FileTitle) + 1, 15)), CHARn
End If
Close #2
Next
Close #1
MsgBox vbCrLf + vbCrLf + Space(20) + "Finished !" + Space(30) + vbCrLf + vbCrLf, vbInformation, "Decompile !"
Err:
Label8.Visible = False
End Sub

Private Sub OrgText_Change()
NewText.MaxLength = Len(OrgText.Text)
End Sub

Private Sub VScrImages_Change()
VScrImages_Scroll
End Sub

Private Sub VScrImages_Scroll()
On Error Resume Next
ExtractedFiles.ListIndex = VScrImages.Value
PicImages.Picture = LoadPicture(App.Path + "\Temp\" + ExtractedFiles.List(VScrImages.Value))
End Sub

Private Sub VScrSentences_Change()
VScrSentences_Scroll
End Sub

Private Sub VScrSentences_Scroll()
On Error GoTo Err
ExtractedFilesSentences.ListIndex = VScrSentences.Value
Dim X As String
Open App.Path + "\Temp\" + ExtractedFilesSentences.List(VScrSentences.Value) For Input As #1
Input #1, X
Close #1
TxtSentences.Text = X
TxtSentences.MaxLength = Len(X)
Err:
End Sub
