VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Saver Setup"
   ClientHeight    =   2535
   ClientLeft      =   225
   ClientTop       =   1530
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame fmScn 
      Caption         =   "Screen Motion"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2415
      Begin MSComctlLib.Slider sdScreen 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   5
         TextPosition    =   1
      End
      Begin VB.Label lblFastScr 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         Height          =   195
         Left            =   1875
         TabIndex        =   6
         Top             =   360
         Width           =   300
      End
      Begin VB.Label lblSlowScr 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   345
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fmTxt 
      Caption         =   "Text Options"
      Height          =   1575
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      Begin VB.CommandButton cmdText 
         Caption         =   "Format Text..."
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.Slider sdText 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   1
         Max             =   50
         SelStart        =   1
         TickFrequency   =   10
         Value           =   1
         TextPosition    =   1
      End
      Begin VB.Label lblSlowTxt 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblFastTxt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         Height          =   195
         Left            =   1875
         TabIndex        =   8
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
    ' Show the about form
    frmAbout.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    ' Exit the form without saving changes
    End
End Sub

Private Sub cmdOkay_Click()
    ' Exit the form and save the settings
    SaveSettings
    End
End Sub

Private Sub cmdText_Click()
    ' Prompt the common dialog box for a new font
    CDialog.FontBold = FtBold
    CDialog.FontItalic = FtItalic
    CDialog.FontUnderline = FtUnderline
    CDialog.FontSize = FtSize
    CDialog.FontName = FtName
    CDialog.FLAGS = cdlCFScreenFonts
    CDialog.ShowFont
    FtBold = CDialog.FontBold
    FtItalic = CDialog.FontItalic
    FtUnderline = CDialog.FontUnderline
    FtSize = CDialog.FontSize
    FtName = CDialog.FontName
End Sub

Private Sub Form_Load()
    ' Initialize slider values
    sdScreen.Value = MovLen
    sdText.Value = sdText.Max - TxtInt
End Sub

Private Sub sdScreen_Change()
    MovLen = sdScreen.Value
End Sub

Private Sub sdText_Change()
    TxtInt = sdText.Max - sdText.Value
End Sub
