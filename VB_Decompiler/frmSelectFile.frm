VERSION 5.00
Begin VB.Form frmSelectFile 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Select Program to Decompile"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      Picture         =   "frmSelectFile.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
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
         TabIndex        =   8
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4800
      Picture         =   "frmSelectFile.frx":13DA
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "O.k"
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
         TabIndex        =   6
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.FileListBox FilesList 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.DirListBox DirList 
      Appearance      =   0  'Flat
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.DriveListBox DriveList 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
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
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   6120
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label CpFileName 
      BackStyle       =   0  'Transparent
      Caption         =   "Not selected"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select EXE or DLL or OCX file"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   160
      Width           =   2580
   End
End
Attribute VB_Name = "frmSelectFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DirList_Change()
FilesList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
On Error GoTo Err
DirList.Path = DriveList.Drive
Exit Sub
Err: MsgBox "Drive not supported !", vbExclamation, "Error"
End Sub

Private Sub FilesList_Click()
If Right(DirList.Path, 1) = "\" Then
FilePath = DirList.Path
Else
FilePath = DirList.Path + "\"
End If
CpFileName.Caption = FilePath + FilesList.List(FilesList.ListIndex)
End Sub

Private Sub Form_Load()
FilesList.Path = App.Path + "\"
DirList.Path = App.Path + "\"
End Sub

Private Sub Label3_Click()
If FilesList.ListIndex = -1 Then Exit Sub
If Not CpFileName.Caption = "Not selected" Then
Label2.Visible = True
DoEvents
Open App.Path + "\OrginalFile.dat" For Output As #1
Write #1, Left(FilesList.List(FilesList.ListIndex), Len(FilesList.List(FilesList.ListIndex)) - 4)
Close #1
FileCopy CpFileName.Caption, App.Path + "\" + FilesList.List(FilesList.ListIndex)
FileName = CpFileName.Caption
frmExtractingData.Show
Unload Me
Else
MsgBox "Select Program to Decompile !", vbInformation, "Wait !"
End If
End Sub

Private Sub Label4_Click()
End
End Sub

