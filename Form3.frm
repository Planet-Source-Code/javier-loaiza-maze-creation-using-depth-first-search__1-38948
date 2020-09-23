VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   6555
   ClientLeft      =   3060
   ClientTop       =   2115
   ClientWidth     =   8610
   LinkTopic       =   "Form3"
   ScaleHeight     =   6555
   ScaleWidth      =   8610
   Begin VB.CommandButton Command1 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "Mael"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picVD 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   10
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   5943
      Left            =   360
      ScaleHeight     =   39.596
      ScaleMode       =   0  'User
      ScaleWidth      =   37.979
      TabIndex        =   6
      Top             =   360
      Width           =   5700
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   178
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   0
         Width           =   178
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Controles"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   6240
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "W = Arriba"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "D = Derecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "S = Abajo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "A = Izkierda"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j As Integer

Const ancho As Integer = 20
Const largo As Integer = 20

Dim r, p As Integer
Dim temp(0 To 3) As Integer
Dim x, y, z, a, b, c, d, e As Integer
Dim cord1(0 To 1) As Integer
Dim cord2(0 To 1) As Integer

Dim matriz1(0 To (ancho - 1), 0 To (largo - 1)) As Integer
Dim matriz2(0 To (ancho * largo - 1), 0 To (ancho * largo - 1)) As Integer
Dim matriz3(0 To (ancho * largo - 1), 0 To (ancho * largo - 1)) As Integer

Dim lab(0 To (ancho * largo - 1)) As Integer
Dim vert(0 To (ancho * largo - 1)) As Integer
Dim cosa As String
Dim key As Integer


Public Sub Command1_Click()
picVD.Cls
For r = 0 To (ancho * largo - 1)
    For p = 0 To (ancho * largo - 1)
        matriz3(r, p) = 0
    Next p
Next r

Randomize
d = Int(((largo) - 0) * Rnd + 0)

picVD.Line (d * 10 + 5, 0)-(d * 10 + 5, 5)
Shape1.Left = d * 10 + 2
Shape1.Top = 1


Randomize
c = Int(((largo) - 0) * Rnd + 0)
e = c
picVD.Line (c * 10 + 5, ancho * 10)-(c * 10 + 5, ancho * 10 - 5)
x = 0
Label1.Caption = e

'Randomize
'c = Int(((390) - 0) * Rnd + 0)
c = 0



lab(x) = c
siguiente

For r = 0 To (ancho * largo - 1)
    If lab(r) = d Then
        d = lab(r)
        Exit For
    End If
Next r


Label1.Caption = e

picVD.SetFocus

End Sub

Private Sub Command2_Click()

Unload Me
Form1.Show

End Sub


Private Sub Form_Load()

picVD.ScaleHeight = ancho * 10
picVD.ScaleWidth = largo * 10

x = 0

For i = 0 To (ancho - 1)
    For j = 0 To (largo - 1)
        matriz1(i, j) = x
        x = x + 1
    Next j
Next i
    
For i = 0 To (ancho - 1)
    For j = 0 To (largo - 1)
    
        'local
        x = i * (largo) + j
             
        
        'arriba
        y = (i - 1) * (largo) + j
        If y >= 0 And y <= (largo * ancho - 1) Then
            matriz2(x, y) = 1
        End If
        
        
        If (j - 1) < 0 Then
        Else
        'izkierda
        y = i * (largo) + (j - 1)
        If y >= 0 And y <= (largo * ancho - 1) Then
            matriz2(x, y) = 1
        End If
        End If
        
        
        If (j + 1) > (largo - 1) Then
        Else
        'derecha
        y = i * (largo) + (j + 1)
        If y >= 0 And y <= (largo * ancho - 1) Then
            matriz2(x, y) = 1
        End If
        End If
        
        'abajo
        y = (i + 1) * (largo) + j
        If y >= 0 And y <= (largo * ancho - 1) Then
            matriz2(x, y) = 1
        End If
        
                
    Next j
Next i


End Sub

Public Function posibles()

z = 0
For i = 0 To 3
    temp(i) = 0
Next i

If x >= 0 Then
    
    For i = 0 To (ancho * largo - 1)
        If matriz2(lab(x), i) = 1 And vert(i) = 0 Then
            temp(z) = i
            z = z + 1
    End If
    Next i
    vertice
Else
    Exit Function
End If

End Function


Public Function vertice()

If temp(0) <> 0 Then
    Randomize
    y = Int(((z) - 0) * Rnd + 0)
    x = x + 1
    
    lab(x) = temp(y)
    matriz3(lab(x - 1), lab(x)) = 1
    matriz3(lab(x), lab(x - 1)) = 1
    'MsgBox (lab(x))
    vert(temp(y)) = 1
    
    For a = 0 To (ancho - 1)
        For b = 0 To (largo - 1)
            If lab(x - 1) = matriz1(a, b) Then
                cord1(0) = a
                cord1(1) = b
            End If
            If lab(x) = matriz1(a, b) Then
                cord2(0) = a
                cord2(1) = b
            End If
            
        Next b
    Next a
    picVD.Line (cord1(1) * 10 + 5, cord1(0) * 10 + 5)-(cord2(1) * 10 + 5, cord2(0) * 10 + 5)
     
    
Else
    x = x - 1
    If x < 0 Then
        Exit Function
    End If
    
    posibles
    
End If

End Function

Public Function siguiente()

For i = 0 To (ancho * largo - 1)
    vert(i) = 0
Next i

vert(c) = 1

Do While x < (largo * ancho - 1) And x >= 0
    posibles
    z = 0
Loop

End Function


Private Sub picVD_KeyPress(KeyAscii As Integer)


'MsgBox ("arriba")

If KeyAscii = 119 Then
    If (d - largo) >= 0 Then
        If matriz3(d, d - largo) = 1 Then
            Shape1.Top = Shape1.Top - 10
            d = d - largo
            If d = e + (largo * ancho - largo) Then
                MsgBox ("Encontraste la salida")
                Command1.SetFocus
            End If
        End If
    End If
End If


'MsgBox ("izkierda")

If KeyAscii = 97 Then
    If (d - 1) >= 0 Then
            If matriz3(d, d - 1) = 1 Then
            Shape1.Left = Shape1.Left - 10
            d = d - 1
            If d = e + (largo * ancho - largo) Then
                MsgBox ("Encontraste la salida")
                Command1.SetFocus
            End If
        End If
    End If
End If


'MsgBox ("abajo")

If KeyAscii = 115 Then
        If d + largo <= (largo * ancho - 1) Then
            If matriz3(d, d + largo) = 1 Then
                Shape1.Top = Shape1.Top + 10
                d = d + largo
                If d = e + (largo * ancho - largo) Then
                    MsgBox ("Encontraste la salida")
                    Command1.SetFocus
                End If
            End If
    End If
End If

'MsgBox ("derecha")

If KeyAscii = 100 Then
    If d < (largo * ancho - 1) Then
        If matriz3(d, d + 1) = 1 Then
            Shape1.Left = Shape1.Left + 10
            d = d + 1
            If d = e + (largo * ancho - largo) Then
                MsgBox ("Encontraste la salida")
                Command1.SetFocus
            End If
        End If
    End If
End If

Label6.Caption = d

End Sub


