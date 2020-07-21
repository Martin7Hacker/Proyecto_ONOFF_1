VERSION 5.00
Begin VB.Form frmAcercade 
   Caption         =   "Acerca de"
   ClientHeight    =   915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   915
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Martin Grasso"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acerca de Encender  5 Led´s con Arduino "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2910
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
