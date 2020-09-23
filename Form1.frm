VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   5880
   ClientTop       =   3465
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   3150
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Experto"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Medio"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Principiante"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Dificultad"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Jugar"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Menu menu1 
      Caption         =   "Opciones"
      Begin VB.Menu Datos 
         Caption         =   "Datos"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim largo As Integer
Dim ancho As Integer

Public Sub Command1_Click()



If Option1.Value = True Then
    Unload Me
    Form2.Show
End If

If Option2.Value = True Then
    Unload Me
    Form3.Show
End If

If Option3.Value = True Then
    Unload Me
    Form4.Show
End If

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Datos_Click()
'MsgBox ("Miembros : "vbcrlf "Javier Loaiza" vbcrlf "Oscar")

End Sub

