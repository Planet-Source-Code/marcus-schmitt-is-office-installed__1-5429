VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function FileExists(sFileName$) As Boolean


On Error Resume Next
FileExists = IIf(Dir(Trim(sFileName)) <> "", True, False)
End Function

Public Function IsAppPresent(strSubKey$, strValueName$) As Boolean

IsAppPresent = CBool(Len(GetRegString(HKEY_CLASSES_ROOT, strSubKey, strValueName)))
End Function

Private Sub Command1_Click()
    Label1.Caption = "Access " & IsAppPresent("Access.Database\CurVer", "")
    Label2.Caption = "Excel " & IsAppPresent("Excel.Sheet\CurVer", "")
    Label3.Caption = "PowerPoint " & IsAppPresent("PowerPoint.Slide\CurVer", "")
    Label4.Caption = "Word " & IsAppPresent("Word.Document\CurVer", "")
End Sub

