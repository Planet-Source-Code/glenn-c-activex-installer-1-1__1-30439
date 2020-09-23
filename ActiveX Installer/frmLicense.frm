VERSION 5.00
Begin VB.Form frmLicense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License Agreement"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1020
      TabIndex        =   3
      Top             =   3270
      Width           =   975
   End
   Begin VB.CommandButton cmdAgree 
      Caption         =   "&Agree"
      Default         =   -1  'True
      Height          =   360
      Left            =   2565
      TabIndex        =   2
      Top             =   3270
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2685
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   4350
   End
   Begin VB.Label Label1 
      Caption         =   "You must agree to this license in order to install this software."
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   105
      Width           =   4275
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form is used to display the license agreement... Duh

Public Canceled As Boolean

Private Sub cmdAgree_Click()
    Canceled = False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Canceled = True
    Unload Me
End Sub
