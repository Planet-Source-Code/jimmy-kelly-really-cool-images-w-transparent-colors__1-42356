VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg2 
      Left            =   105
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   16711935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pick Transparent Color"
      Height          =   1170
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   1275
   End
   Begin VB.PictureBox Buffer 
      Height          =   5685
      Left            =   4725
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   214
      TabIndex        =   3
      Top             =   105
      Width           =   3270
   End
   Begin VB.PictureBox Pic 
      Height          =   5685
      Left            =   1365
      Picture         =   "Form1.frx":7984
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   214
      TabIndex        =   2
      Top             =   105
      Width           =   3270
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy"
      Height          =   1170
      Left            =   0
      TabIndex        =   1
      Top             =   1260
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   735
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "dlg1"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pick Color"
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function _
TransparentBlt Lib "msimg32" (ByVal _
hDestDC As Long, ByVal X As Long, ByVal _
Y As Long, ByVal nWidth As Long, ByVal _
nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal nWidthSrc As Long, ByVal _
nHeightSrc As Long, ByVal crTrans As _
Long) As Long


Private Sub Command1_Click()

dlg1.ShowColor
Pic.ForeColor = dlg1.Color

End Sub

Private Sub Command2_Click()
Buffer.Cls
Dim lngrval As Integer
   lngrval = TransparentBlt(Buffer.hDC, 0, 0, Buffer.ScaleWidth, Buffer.ScaleHeight, Pic.hDC, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, dlg2.Color)
End Sub

Private Sub Command3_Click()

dlg2.ShowColor

End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Pic.DrawWidth = 32
If Button = 1 Then Pic.PSet (X, Y)

End Sub
