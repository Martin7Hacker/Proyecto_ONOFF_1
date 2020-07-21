VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPrograma 
   Caption         =   "LED´s con Arduino UNO"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4875
   Icon            =   "led.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      Caption         =   "USB"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PIN 6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PIN 5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PIN 4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PIN 3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PIN 2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSCommLib.MSComm USB 
      Left            =   720
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APAGADO"
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
      TabIndex        =   9
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APAGADO"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APAGADO"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APAGADO"
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
      TabIndex        =   3
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APAGADO"
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
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' leds con Arduino UNO
' puede ser con otros Tambien
' Autor Martin grasso.
' recuerde cargar el Archivo al Arduino
' antes de utilizar el programa
' y verificar el puerto de comunciacion
'
'Puerto donde esta el Arduino
Public puerto As Integer
Dim contador As Integer

Private Sub Command1_Click()
 boton_LED Command1, Label1, 1, 0
End Sub
'
Private Sub Command2_Click()
boton_LED Command2, Label2, 3, 2
End Sub
'
Private Sub Command3_Click()
boton_LED Command3, Label3, 5, 4
End Sub
'
Private Sub Command4_Click()
boton_LED Command4, Label4, 7, 6
End Sub
'
Private Sub Command5_Click()
boton_LED Command5, Label5, 9, 8
End Sub
'
'comandos para apagar y encender el led en Arduino
'recomendamos arduino UNO
'
Private Sub boton_LED(ByVal boton As CommandButton, ByVal Label _
As Label, ByVal encendido As String, apagado As String)
Form_Load
If Label.Caption = "APAGADO" Then
         Label.Caption = "ENCENDIDO"
         Label.ForeColor = vbRed
         USB.Output = apagado     'Mandamos un 0 para encender el led en el pin 2
         contador = contador + 1
    Else
        Label.Caption = "APAGADO"
        Label.ForeColor = vbBlack
        USB.Output = encendido      'Mandamos un 1 para apagar el led en el pin 2
        contador = contador - 1
    End If
End Sub

Private Sub Command7_Click()
frmAcercade.Show 1
End Sub

Private Sub Command8_Click()
frmpuerto.Show 1
End Sub

'
'
'cuando se carge el programa
'configurar el puerto de salida serial
'del dispositivo usb
'
'
Private Sub Form_Load()
cargar_Driver
On Error GoTo nose
With USB
    .RThreshold = 1
    .InputLen = 1
    .Settings = "9600" 'velocidad en baudios
    .CommPort = puerto          'numero de puerto utilizado // definalo en su arduino es el puerto donde
                           'el arduino esta conectado
                           
    .InBufferSize = 1  'Tamano del Buffer de entrada
    .InputLen = 1      'cantidad de datos a leer
    .DTREnable = False 'Deshabilitar el Threshold para TR
    .PortOpen = True   ' Abre el puerto"
End With
nose:
End Sub

Private Sub Form_Unload(Cancel As Integer)
If contador >= 1 Then
Cancel = True
Select Case MsgBox("¿Quieres Salir y Apagar todas las LED´s?", vbYesNo, "Advertencia")
Case vbYes
     With USB
     .Output = "1"
     .Output = "3"
     .Output = "5"
     .Output = "7"
     .Output = "9"
     Cancel = False
     End With
End Select
End If
End Sub

Private Sub cargar_Driver()
Dim driv As String
On Error GoTo nose
Open App.Path & "\Drivers.hex" For Input As 1
 Do While Not EOF(1)
  Line Input #1, driv
  frmPrograma.puerto = (driv)
 Loop
 Close #1
nose:
End Sub
