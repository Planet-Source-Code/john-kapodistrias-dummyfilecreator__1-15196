VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "          [ DummyFileCreator ]"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Custom Buffer"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Random Buffer"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   173
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   375
      Left            =   593
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   3120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   3120
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   1560
      Y2              =   3840
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   3120
      Y1              =   1560
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3120
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label7 
      Caption         =   "extension"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "name"
      Height          =   195
      Left            =   1440
      TabIndex        =   12
      Top             =   3120
      Width           =   390
   End
   Begin VB.Label Label5 
      Caption         =   "Name of files"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Custom Buffer Contents"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Fill with"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Options"
      Height          =   255
      Left            =   1133
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number of files to create"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   413
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CreateDF Text1.Text, App.Path
End Sub

Private Sub CreateDF(NumOfFiles As Integer, Path As String)
On Error Resume Next
Dim Buf As String
If Option1.Value = True Then
    Buf = Rnd * 123456789
ElseIf Option2.Value = True Then
    Buf = Text2.Text
End If
Dim NM As String
Dim EX As String
If Text3.Text = "" Then
    NM = "DummyFile"
    EX = "dat"
Else
    NM = Text3.Text
    EX = Text4.Text
End If
Dim I As Integer
Dim T As Integer
I = 1
Dim X As Integer
X = NumOfFiles
Dim FPath As String
 For T = 1 To X Step 1
    I = I + 1
    FPath = Path & "\" & NM & I & "." & EX
    Open FPath For Output As #1
    Print #1, Buf
    Close #1
 Next T
End Sub
Public Sub PaintBlue(frmForm As Form)
Dim ctlControl As Object
    On Error Resume Next
    For Each ctlControl In frmForm.Controls
        If TypeOf ctlControl Is Label Then
        ctlControl.ForeColor = vbHighlight
        End If
        DoEvents
    Next ctlControl
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call PaintBlue(Me)
Me.Height = 1755
End Sub

Private Sub Label2_Click()
Select Case Me.Height
 Case 1755
    Label2.BorderStyle = 1
    Me.Height = 4320
 Case 4320
    Label2.BorderStyle = 0
    Me.Height = 1755
End Select
End Sub


