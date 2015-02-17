VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3555
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3450
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   120
         Top             =   240
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "(c) 2005 André Martins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   1
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versão BETA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5325
         TabIndex        =   2
         Top             =   2700
         Width           =   1530
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "for Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4965
         TabIndex        =   3
         Top             =   2340
         Width           =   1890
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Down Blocks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   4065
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "AndreSoft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   4
         Top             =   600
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmJogo.Show
    Unload Me
End Sub

Private Sub Frame1_Click()
    frmJogo.Show
    Unload Me
End Sub

Private Sub imgLogo_Click()
    frmJogo.Show
    Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
    frmJogo.Show
    Unload Me
End Sub

Private Sub lblCopyright_Click()
    frmJogo.Show
    Unload Me
End Sub

Private Sub lblPlatform_Click()
    frmJogo.Show
    Unload Me
End Sub

Private Sub lblProductName_Click()
    frmJogo.Show
    Unload Me
End Sub

Private Sub lblVersion_Click()
    frmJogo.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
    frmJogo.Show
    Unload Me
End Sub
