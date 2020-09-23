VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base64 Encoding/Decoding"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   6375
      Begin VB.CommandButton Command4 
         Caption         =   "Decode"
         Height          =   495
         Left            =   2640
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Input Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Output Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6375
      Begin VB.CommandButton Command3 
         Caption         =   "Encode"
         Height          =   495
         Left            =   2640
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Input Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Output Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6375
      Begin VB.CommandButton Command2 
         Caption         =   "Decode"
         Height          =   495
         Left            =   2640
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6375
      Begin VB.TextBox Text1 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   240
         Width           =   6135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Encode"
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   3120
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip3 
      Height          =   4335
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "String"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "String"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Encoding"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Decoding"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As New b64
Private Sub Command1_Click()
Text1.Text = b.encode(Text1.Text)
End Sub

Private Sub Command2_Click()
Text2.Text = b.decode(Text2.Text)
End Sub
Private Sub Command3_Click()
If Trim(Text3.Text) <> "" And Trim(Text4.Text) <> "" Then
b.encodefile Text4.Text, Text3.Text
MsgBox "Done", , App.Title
Exit Sub
End If
MsgBox "Not Done", , App.Title
End Sub
Private Sub Command4_Click()
If Trim(Text5.Text) <> "" And Trim(Text6.Text) <> "" Then
b.decodefile Text6.Text, Text5.Text
MsgBox "Done", , App.Title
Exit Sub
End If
MsgBox "Not Done", , App.Title
End Sub

Private Sub Form_Load()
App.Title = "Base64 Encoder/Decoder"
TabStrip2.Visible = True
TabStrip3.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame3.Visible = True
End Sub

Private Sub TabStrip1_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
If TabStrip1.Tabs(1).Selected Then
TabStrip2.Visible = True
TabStrip3.Visible = False
Call TabStrip2_Click
Else
TabStrip3.Visible = True
TabStrip2.Visible = False
Call TabStrip3_Click
End If
End Sub

Private Sub TabStrip2_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
If TabStrip2.Tabs(1).Selected Then
Frame3.Visible = True
Else
Frame1.Visible = True
End If
End Sub
Private Sub TabStrip3_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
If TabStrip3.Tabs(1).Selected Then
Frame4.Visible = True
Else
Frame2.Visible = True
End If
End Sub

Private Sub Text3_Click()
cdl1.ShowOpen
Text3.Text = cdl1.FileName
End Sub
Private Sub Text4_Click()
cdl1.ShowOpen
Text4.Text = cdl1.FileName
End Sub
Private Sub Text5_Click()
cdl1.ShowOpen
Text5.Text = cdl1.FileName
End Sub
Private Sub Text6_Click()
cdl1.ShowOpen
Text6.Text = cdl1.FileName
End Sub

