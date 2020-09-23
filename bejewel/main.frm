VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bejewel"
   ClientHeight    =   9315
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11805
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   621
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   787
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Dibujar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9360
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mesclar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   9360
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   9240
      Top             =   7320
      _ExtentX        =   6350
      _ExtentY        =   3175
      _Version        =   393216
      Rows            =   2
      Cols            =   4
      Picture         =   "main.frx":1982
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   9030
      Left            =   120
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   120
      Width           =   9030
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   10320
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntaje:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9360
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu njuego 
      Caption         =   "Nuevo Juego"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private la_matrix(9, 9) As Byte
Private Only_Once As Boolean
Private unx As Byte
Private uny As Byte
Private unCell As Byte
Private enSel As Boolean

Private Sub elMesclar()
 elComponer
 elDibujar
End Sub

Private Sub Command1_Click()
 elMesclar
End Sub

Private Sub Command2_Click()
 elDibujar
End Sub

Private Sub Form_Activate()
 If Only_Once = False Then
  Only_Once = True
  elDibujar
 End If
End Sub

Private Sub Form_Load()
 Only_Once = False
 enSel = False
 elComponer
End Sub

Private Sub elDibujar()
 Dim x As Byte
 Dim y As Byte
  y = 0
  Do While y <= 9
   x = 0
   Do While x <= 9
    Picture1.PaintPicture PictureClip1.GraphicCell(la_matrix(y, x)), x * 60, y * 60
    x = x + 1
   Loop
   y = y + 1
  Loop
End Sub

Private Sub elComponer()
 Dim x As Byte
 Dim y As Byte
 Dim elRand As Byte
 Dim elOK As Boolean
 Randomize
 y = 0
 Do While y <= 9
  x = 0
  Do While x <= 9
   elOK = False
   Do While elOK = False
    elOK = True
    elRand = 7 * Rnd()
    If x > 1 Then
     If la_matrix(y, x - 1) = elRand And la_matrix(y, x - 2) = elRand Then elOK = False
    End If
    If y > 1 Then
     If la_matrix(y - 1, x) = elRand And la_matrix(y - 2, x) = elRand Then elOK = False
    End If
   Loop
   la_matrix(y, x) = elRand
   x = x + 1
  Loop
  y = y + 1
 Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub

Private Sub njuego_Click()
 elMesclar
 Label2.Caption = 0
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim otroY As Byte
 Dim otroX As Byte
 If enSel = False Then
  Picture1.Line (unx * 60, uny * 60)-((unx * 60) + 59, (uny * 60) + 59), vbWhite, B
  uny = Int(y / 60)
  unx = Int(x / 60)
  unCell = la_matrix(uny, unx)
  Picture1.Line (unx * 60, uny * 60)-((unx * 60) + 59, (uny * 60) + 59), vbGreen, B
  enSel = True
 Else
  otroY = Int(y / 60)
  otroX = Int(x / 60)
  If (otroY = uny And (otroX = unx + 1 Or otroX = unx - 1) And la_matrix(uny, unx) <> la_matrix(uny, otroX)) Or (otroX = unx And (otroY = uny + 1 Or otroY = uny - 1) And la_matrix(otroY, unx) <> la_matrix(uny, unx)) Then
    Verificar_Coincidencias otroX, otroY
  Else
   Picture1.Line (unx * 60, uny * 60)-((unx * 60) + 59, (uny * 60) + 59), vbWhite, B
   enSel = False
  End If
   
 End If
End Sub

Private Sub Verificar_Coincidencias(ByVal elX As Byte, ByVal elY As Byte)
 Dim n As Byte
 Dim n2 As Byte
 Dim x As Byte
 Dim y As Byte
 Dim HayCambios As Boolean
 Dim Abortar As Boolean
 Dim yaHizo As Integer
 
 HayCambios = False
 yaHizo = 0
 
 Picture1.Line (unx * 60, uny * 60)-((unx * 60) + 59, (uny * 60) + 59), vbWhite, B
 enSel = False
 Animar_Cambio unx, uny, elX, elY
  
 Abortar = True
 Do While Abortar = True

 Abortar = False
 y = 0
 Do While y <= 9
  x = 0
  n = 0
  n2 = 0
  Do While x <= 9
   If la_matrix(y, x) = la_matrix(y, n) Then
    n2 = x
   Else
    If n2 - n >= 2 Then
     HayCambios = True
     Procesar1 n, n2, y, yaHizo
     yaHizo = yaHizo + 1
     n = 0
     n2 = 0
     y = 0
     Exit Do
    Else
     n = x
     n2 = x
    End If
   End If
   x = x + 1
  Loop
  
  If n2 - n >= 2 Then
   HayCambios = True
   Procesar1 n, n2, y, yaHizo
   y = 0
   Exit Do
  End If
  
  y = y + 1
 Loop
  
 x = 0
 Do While x <= 9 And Abortar = False
  y = 0
  n = 0
  n2 = 0
  Do While y <= 9 And Abortar = False
   If la_matrix(y, x) = la_matrix(n, x) Then
    n2 = y
   Else
    If n2 - n >= 2 Then
     HayCambios = True
     Procesar2 n, n2, x, yaHizo
     yaHizo = yaHizo + 1
       n = 0
       n2 = 0
     Abortar = True
     Exit Do
    Else
     n = y
     n2 = y
    End If
   End If
   y = y + 1
  Loop

    If n2 - n >= 2 Then
     HayCambios = True
     Procesar2 n, n2, x, yaHizo
     Abortar = True
     Exit Do
    End If

  x = x + 1
 Loop
 
 Loop
  
 If HayCambios = False Then
  Animar_Cambio unx, uny, elX, elY
 End If
 
End Sub

Private Sub Procesar1(ByVal n As Byte, ByVal n2 As Byte, ByVal y As Byte, ByVal elYa As Integer)
 Dim j As Integer
 Dim g As Byte
 Dim vector(9) As Byte
 Dim f As Double
 Dim elPunt As Integer
 Randomize
 j = 0
 Do While j <= 9
  vector(j) = 7 * Rnd()
  j = j + 1
 Loop
 j = 0
 Do While j < 60
  g = 0
  Do While g <= n2 - n
   Picture1.Line (((n + g) * 60) + j, (y * 60) + j)-((((n + g) * 60) + 60) - j, ((y * 60) + 60) - j), vbWhite, B
   g = g + 1
  Loop
  Pause 0.004
  j = j + 1
 Loop
 
 If n2 - n = 2 Then elPunt = 10
 If n2 - n = 3 Then elPunt = 15
 If n2 - n = 4 Then elPunt = 25
 
 If elYa > 0 Then elPunt = elPunt * elYa
 
 Label2.Caption = Label2.Caption + elPunt
 
 f = 0
 Do While f <= 60
  j = y - 1
  Do While j >= -1
   If j = -1 Then
    g = n
    Do While g <= n2
     Picture1.PaintPicture PictureClip1.GraphicCell(vector(g)), g * 60, (j * 60) + f
     g = g + 1
    Loop
   Else
    g = n
    Do While g <= n2
     Picture1.PaintPicture PictureClip1.GraphicCell(la_matrix(j, g)), g * 60, (j * 60) + f
     g = g + 1
    Loop
   End If
   j = j - 1
  Loop
  f = f + 1
 Loop
 j = y
 Do While j >= 0
  g = n
  Do While g <= n2
   If j > 0 Then
    la_matrix(j, g) = la_matrix(j - 1, g)
   Else
    la_matrix(j, g) = vector(g)
   End If
   g = g + 1
  Loop
  j = j - 1
 Loop
End Sub

Private Sub Procesar2(ByVal n As Byte, ByVal n2 As Byte, ByVal x As Byte, ByVal elYa As Integer)
 Dim j As Integer
 Dim g As Byte
 Dim vector(9) As Byte
 Dim f As Double
 Dim y As Integer
 Dim elPunt As Integer
 Randomize
 j = 0
 Do While j <= 9
  vector(j) = 7 * Rnd()
  j = j + 1
 Loop
 j = 0
 Do While j < 60
  g = 0
  Do While g <= n2 - n
   Picture1.Line ((x * 60) + j, ((n + g) * 60) + j)-(((x * 60) + 60) - j, (((n + g) * 60) + 60) - j), vbWhite, B
   g = g + 1
  Loop
  Pause 0.004
  j = j + 1
 Loop
 
 If n2 - n = 2 Then elPunt = 10
 If n2 - n = 3 Then elPunt = 15
 If n2 - n = 4 Then elPunt = 25
 
 If elYa > 0 Then elPunt = elPunt * elYa
 
 Label2.Caption = Label2.Caption + elPunt
 
 f = 0
 Do While f <= ((n2 - n) + 1) * 60
  j = n - 1
  g = 0
  Do While j >= -((n2 - n) + 1)
   If j >= 0 Then
    Picture1.PaintPicture PictureClip1.GraphicCell(la_matrix(j, x)), x * 60, (j * 60) + f
   Else
    Picture1.PaintPicture PictureClip1.GraphicCell(vector(g)), x * 60, (j * 60) + f
    g = g + 1
   End If
   j = j - 1
  Loop
  f = f + 1
 Loop
 
 j = n2
 y = n - 1
 g = 0
 
 Do While j >= 0
  If y < 0 Then
   la_matrix(j, x) = vector(g)
   g = g + 1
  Else
   la_matrix(j, x) = la_matrix(y, x)
  End If
  y = y - 1
  j = j - 1
 Loop
End Sub

Private Sub Animar_Cambio(ByVal x1 As Byte, ByVal y1 As Byte, ByVal x2 As Byte, ByVal y2 As Byte)
 Dim n As Double
 Dim n2 As Double
 Dim elG As Integer
 If y1 = y2 Then
  If x1 < x2 Then
   elG = -1
  Else
   elG = 1
  End If
  n = x1 * 60
  n2 = x2 * 60
  Do While n <> (x2 * 60) - elG
   Picture1.PaintPicture PictureClip1.GraphicCell(la_matrix(y1, x1)), n, y1 * 60
   Picture1.PaintPicture PictureClip1.GraphicCell(la_matrix(y2, x2)), n2, y1 * 60
   If x1 < x2 Then
    n = n + 1
    n2 = n2 - 1
   Else
    n = n - 1
    n2 = n2 + 1
   End If
   Pause 0.004
  Loop
 Else
  If y1 < y2 Then
   elG = -1
  Else
   elG = 1
  End If
  n = y1 * 60
  n2 = y2 * 60
  Do While n <> (y2 * 60) - elG
   Picture1.PaintPicture PictureClip1.GraphicCell(la_matrix(y1, x1)), x1 * 60, n
   Picture1.PaintPicture PictureClip1.GraphicCell(la_matrix(y2, x2)), x1 * 60, n2
   If y1 < y2 Then
    n = n + 1
    n2 = n2 - 1
   Else
    n = n - 1
    n2 = n2 + 1
   End If
   Pause 0.004
  Loop
 End If
 n = la_matrix(y1, x1)
 la_matrix(y1, x1) = la_matrix(y2, x2)
 la_matrix(y2, x2) = n
End Sub
