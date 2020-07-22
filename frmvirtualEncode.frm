VERSION 5.00
Begin VB.Form frmvirtualEncode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Encode v1.0 free by Martin Hacker"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmvirtualEncode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmduno 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   105
      Width           =   255
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "(F5) &Exit"
      Height          =   495
      Left            =   5400
      TabIndex        =   18
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtdescriptarM 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "_"
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox txtEncripterM 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "_"
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.ComboBox cobfrecuencia2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdAcercade 
      Caption         =   " Donate"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdremoveMasc 
      Caption         =   "(F3) &Remove mask"
      Height          =   495
      Left            =   2450
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdmasEm 
      Caption         =   "&(F2) &Emergency mask"
      Height          =   495
      Left            =   1090
      TabIndex        =   11
      Top             =   1920
      Width           =   1320
   End
   Begin VB.CommandButton cmddecoder 
      Caption         =   "(F4) &decode"
      Height          =   495
      Left            =   3570
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdescriptar 
      Caption         =   "(F1) &encode"
      Height          =   495
      Left            =   100
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdcopiarx 
      Caption         =   "x"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdborrar 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.ComboBox cobfrecuencia 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtdescriptar 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
   End
   Begin VB.TextBox txtEncripter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label lblsoftwareprogramer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software Programmed by: Martin Grasso Castrillo."
      ForeColor       =   &H00800080&
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   3765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual Encode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   420
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   3120
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      FillColor       =   &H00800080&
      Height          =   375
      Left            =   3960
      Top             =   105
      Width           =   2085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "code to decode:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5850
   End
   Begin VB.Label Label2 
      Caption         =   "code to encode:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2850
   End
   Begin VB.Label Label1 
      Caption         =   "frequency: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800080&
      BorderWidth     =   5
      Height          =   495
      Left            =   120
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800080&
      BorderWidth     =   5
      Height          =   495
      Left            =   120
      Top             =   600
      Width           =   5895
   End
   Begin VB.Menu options 
      Caption         =   "options"
      Begin VB.Menu F1 
         Caption         =   "&encode"
         Shortcut        =   {F1}
      End
      Begin VB.Menu F2 
         Caption         =   "&Emergency mask"
         Shortcut        =   {F2}
      End
      Begin VB.Menu F3 
         Caption         =   "Remove mask"
         Shortcut        =   {F3}
      End
      Begin VB.Menu F4 
         Caption         =   "decode"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Donate 
         Caption         =   "&Donate"
         Shortcut        =   {F6}
      End
      Begin VB.Menu help 
         Caption         =   "&help"
         Shortcut        =   {F7}
      End
      Begin VB.Menu F5 
         Caption         =   "&Exit"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "F7 (&?)"
   End
End
Attribute VB_Name = "frmvirtualEncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'                                        *
' Programa para Decodificar texto en     *
' formato Digital                        *
' Autor Martin Grasso 2017               *
'                                        *
'*****************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim datos As New clsEscritar

Private Sub cmdCopiar_Click()
Clipboard.SetText ("")
Clipboard.SetText (txtEncripter.Text)
Clipboard.SetText (txtEncripterM.Text)
End Sub

Private Sub ayuda_Click()
AbrirWeb Me, "http://adf.ly/1VJBe9"
End Sub

Private Sub cmdAcercade_Click()
AbrirWeb Me, "http://adf.ly/1TJlyy"
End Sub

Private Sub cmdborrar_Click()
txtEncripterM.Text = ""
txtEncripter.Text = ""
End Sub

Private Sub cmdcopiarx_Click()
txtdescriptar.Text = ""
txtdescriptarM.Text = ""
End Sub

Private Sub cmddecoder_Click()
txtEncripter.Text = datos.funcion_desescriptar(txtdescriptar.Text)
txtEncripterM.Text = datos.funcion_desescriptar(txtdescriptarM.Text)
End Sub

Private Sub cmdescriptar_Click()
txtdescriptar.Text = ""
txtdescriptarM.Text = ""
txtdescriptar.Text = datos.funcion_escriptar(txtEncripter.Text)
txtdescriptarM.Text = datos.funcion_escriptar(txtEncripterM.Text)
End Sub

Private Sub cmdmasEm_Click()
mascara False, True
End Sub

Private Sub cmdremoveMasc_Click()
mascara True, False
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmduno_Click()
cobfrecuencia2.ListIndex = 1
cobfrecuencia.ListIndex = 1
End Sub

Private Sub cobfrecuencia_Change()
cobfrecuencia2.ListIndex = cobfrecuencia.ListIndex
ingresarDatos
End Sub

Private Sub cobfrecuencia_Click()
cobfrecuencia_Change
End Sub

Private Sub Donate_Click()
cmdAcercade_Click
End Sub

Private Sub F1_Click()
cmdescriptar_Click
End Sub

Private Sub F2_Click()
cmdmasEm_Click
End Sub

Private Sub F3_Click()
cmdremoveMasc_Click
End Sub

Private Sub F4_Click()
cmddecoder_Click
End Sub

Private Sub F5_Click()
cmdsalir_Click
End Sub

Private Sub Form_Load()
cobfrecuencia.Clear
frecuencia
End Sub

Private Sub frecuencia()
Dim fx As Integer
For fx = 0 To 100
    cobfrecuencia.AddItem fx & " hz"
    cobfrecuencia2.AddItem fx
Next fx
cobfrecuencia.ListIndex = 1
End Sub

Private Sub ingresarDatos()
datos.datoEncoder = cobfrecuencia2.List(cobfrecuencia.ListIndex)
End Sub

Private Sub mascara(ByVal control1 As Boolean, ByVal control2 As Boolean)
txtEncripter.Visible = control1
txtEncripterM.Visible = control2
'----------
txtdescriptar.Visible = control1
txtdescriptarM.Visible = control2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Select Case MsgBox("Exit the application?", vbYesNo + vbInformation, "Virtual Encode v1.0 free by: Martin Hacker")
       Case (vbYes)
      Cancel = 0
End Select
End Sub

Private Sub help_Click()
ayuda_Click
End Sub

Private Sub lblsoftwareprogramer_Click()
AbrirWeb Me, "http://adf.ly/1TJrkA"
End Sub

Private Sub txtdescriptar_Change()
txtdescriptarM.Text = txtdescriptar.Text
End Sub

Private Sub txtdescriptar_Click()
txtdescriptar_Change
End Sub

Private Sub txtdescriptarM_Change()
txtdescriptar.Text = txtdescriptarM.Text
End Sub

Private Sub txtdescriptarM_Click()
txtdescriptarM_Change
End Sub

Private Sub txtEncripter_Change()
txtEncripterM.Text = txtEncripter.Text
End Sub

Private Sub txtEncripter_Click()
txtEncripter_Change
End Sub

Private Sub txtEncripterM_Change()
txtEncripter.Text = txtEncripterM.Text
End Sub

Private Sub txtEncripterM_Click()
txtEncripter_Change
End Sub
Public Sub AbrirWeb(ByVal control As Form, ByVal web As String)
Dim x As String
    x = ShellExecute(control.hwnd, "Open", web, &O0, &O0, 0)
'donde mipagina.cl colocas la url que quieras
End Sub

