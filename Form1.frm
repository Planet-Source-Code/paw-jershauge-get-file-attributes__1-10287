VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "By paw Jershauge (PBFJ@Hotmail.com)"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "File info"
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   4335
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   15
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   14
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File size"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File date"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "Set Attribute"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox R 
         Caption         =   "Read only"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox H 
         Caption         =   "Hidden"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox A 
         Caption         =   "Archive"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox S 
         Caption         =   "System"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog C1 
      Left            =   4080
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Filepath As String, FileSize As Long, Filedate As String

Private Sub Command1_Click()
C1.ShowOpen
loadfile
End Sub

Function loadfile()
If C1.FileName <> "" Then
GetInfo (C1.FileName)
Text1.Text = C1.FileName
Filepath = Mid(C1.FileName, 1, InStr(1, C1.FileName, C1.FileTitle, vbTextCompare) - 1)
FileSize = FileLen(C1.FileName)
Filedate = FileDateTime(C1.FileName)
Label2(0).Caption = Filepath
Label2(1).Caption = C1.FileTitle
Label2(2).Caption = Filedate
Label2(3).Caption = FileSize & " Byte(s)"
Command2.Enabled = True
Text1.Locked = True
Else
Command2.Enabled = False
Text1.Locked = False
End If
End Function

Function GetInfo(filen As String)
Select Case GetAttr(filen)
Case 1
A.Value = 0
R.Value = 1
H.Value = 0
S.Value = 0
Case 2
A.Value = 0
R.Value = 0
H.Value = 1
S.Value = 0
Case 3
A.Value = 0
R.Value = 1
H.Value = 1
S.Value = 0
Case 4
A.Value = 0
R.Value = 0
H.Value = 0
S.Value = 1
Case 5
A.Value = 0
R.Value = 1
H.Value = 0
S.Value = 1
Case 6
A.Value = 0
R.Value = 0
H.Value = 1
S.Value = 1
Case 7
A.Value = 0
R.Value = 1
H.Value = 1
S.Value = 1
Case 32
A.Value = 1
R.Value = 0
H.Value = 0
S.Value = 0
Case 33
A.Value = 1
R.Value = 1
H.Value = 0
S.Value = 0
Case 34
A.Value = 1
R.Value = 0
H.Value = 1
S.Value = 0
Case 35
A.Value = 1
R.Value = 1
H.Value = 1
S.Value = 0
Case 36
A.Value = 1
R.Value = 0
H.Value = 0
S.Value = 1
Case 38
A.Value = 1
R.Value = 0
H.Value = 1
S.Value = 1
Case 39
A.Value = 1
R.Value = 1
H.Value = 1
S.Value = 1
End Select
End Function

Private Sub Command2_Click()
Dim userset As Long
If A.Value = 1 Then
userset = userset + 32
End If
If R.Value = 1 Then
userset = userset + 1
End If
If H.Value = 1 Then
userset = userset + 2
End If
If S.Value = 1 Then
userset = userset + 4
End If
SetAttr Text1.Text, userset
loadfile
End Sub
