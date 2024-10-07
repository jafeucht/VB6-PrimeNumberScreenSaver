VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ..."
   ClientHeight    =   1935
   ClientLeft      =   1140
   ClientTop       =   1560
   ClientWidth     =   3675
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   135
      TabIndex        =   1
      Top             =   1440
      Width           =   3390
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum code to create a screen saver using Visual Basic This example was created by Igguk and modified by Feuchtersoft."
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblDeco1 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Saver 1.0"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()

    ' Unload and deallocate the about box.
    End

End Sub

