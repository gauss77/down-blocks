VERSION 5.00
Begin VB.Form frmSobre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre Down Blocks"
   ClientHeight    =   3075
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmSobre.frx":0000
   LinkTopic       =   "frmSobre"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2122.42
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Label lblTeste 
      Caption         =   "lblTeste"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()

    'Fecha o formulário
    Unload Me

End Sub

