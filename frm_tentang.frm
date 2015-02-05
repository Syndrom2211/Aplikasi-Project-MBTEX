VERSION 5.00
Begin VB.Form frm_tentang 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tentang"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5505
   Icon            =   "frm_tentang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   3240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2013"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   1440
      Picture         =   "frm_tentang.frx":1CCA
      Top             =   960
      Width           =   2550
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versi 1.0"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAM DATABASE MBTEX"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frm_tentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim countertitle As Integer
Dim title As String
Dim titledance As String

Private Sub Form_Load()
title = "Created by Firdam @ IndonesianCoder"
End Sub

Private Sub Timer1_Timer()
titledance = Left(title, countertitle)
Me.Caption = titledance
countertitle = countertitle + 1
If countertitle >= Len(title) + 3 Then
    countertitle = 1
    titledance = ""
End If
End Sub
