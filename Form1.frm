VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   180
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicChannel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   3570
      Left            =   135
      ScaleHeight     =   3510
      ScaleMode       =   0  'User
      ScaleWidth      =   11393.5
      TabIndex        =   1
      Top             =   1170
      Width           =   13230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   5715
      TabIndex        =   0
      Top             =   315
      Width           =   2220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CD.InitDir = App.Path
    CD.ShowOpen
    Me.MousePointer = vbHourglass
    Me.PicChannel.Cls
    GetWavInfo CD.FileName, Me.PicChannel
    Me.MousePointer = vbNormal
End Sub
