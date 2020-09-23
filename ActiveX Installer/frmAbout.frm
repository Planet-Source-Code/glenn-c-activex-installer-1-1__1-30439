VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   330
      Left            =   1237
      TabIndex        =   4
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "Email:  hardrequest@hotmail.com"
      Height          =   195
      Left            =   510
      TabIndex        =   3
      Top             =   1335
      Width           =   2355
   End
   Begin VB.Label Label3 
      Caption         =   "Version 1.1"
      Height          =   195
      Left            =   2175
      TabIndex        =   2
      Top             =   645
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright © 2002 Glenn Chittenden Jr."
      Height          =   195
      Left            =   322
      TabIndex        =   1
      Top             =   1080
      Width           =   2730
   End
   Begin VB.Label Label1 
      Caption         =   "ActiveX Installer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   405
      TabIndex        =   0
      Top             =   210
      Width           =   2565
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
