VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "VBALPROGBAR6.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar1 
      Height          =   450
      Left            =   240
      TabIndex        =   2
      Top             =   2130
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   794
      Picture         =   "frmMain.frx":0CCA
      BackColor       =   4194368
      ForeColor       =   0
      BorderStyle     =   0
      BarPicture      =   "frmMain.frx":0CE6
      BarPictureMode  =   0
      BackPictureMode =   0
      ShowText        =   -1  'True
      Text            =   "[0% Completado]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmMain.frx":669A
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label LSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0 MBs de 0 MBs"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa actualizará tu cliente a la nueva versión. Para empezar clickea en buscar actualizaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   4200
      Top             =   2880
      Width           =   2385
   End
   Begin VB.Image Image2 
      Height          =   630
      Left            =   240
      Top             =   2880
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&
Dim Directory As String, bDone As Boolean, dError As Boolean, F As Integer
Rem Programado por Shedark

Private Sub Analizar()

    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer, dNum As String
    
    'lEstado.Caption = "Obteniendo datos..."
    Call addConsole("Buscando Actualizaciones...", 255, 255, 255, True, False)
    Call Reproducir_WAV(App.Path & "\Wav\Revision.wav", SND_FILENAME)
    
    iX = Inet1.OpenURL("https://raw.githubusercontent.com/ao-oficial/ao-website/master/argentum-online-client-current-version.txt") 'Host
    tX = LeerInt(App.Path & "\INIT\Update.ini")
    
    DifX = iX - tX
    
    If Not (DifX = 0) Then
    If MsgBox("Se descargarán " & DifX & "actualizaciones, ¿Continuar?", vbYesNo) = vbYes Then
    ProgressBar1.Visible = True

        Call addConsole("Iniciando, se descargarán " & DifX & " actualizaciones.", 200, 200, 200, True, False)   '>> Informacion
        For i = 1 To DifX
            Inet1.AccessType = icUseDefault
            dNum = i + tX
            
            #If BuscarLinks Then 'Buscamos el link en el host (1)
                Inet1.URL = Inet1.OpenURL("https://raw.githubusercontent.com/ao-oficial/ao-website/master/Link" & dNum & ".txt") 'Host
            #Else                'Generamos Link por defecto (0)
                Inet1.URL = "https://raw.githubusercontent.com/ao-oficial/ao-website/master/Parche" & dNum & ".zip" 'Host
            #End If
            
            Directory = App.Path & "\INIT\Parche" & dNum & ".zip"
            bDone = False
            dError = False
            
            'lURL.Caption = Inet1.URL
            'lName.Caption = "Parche" & dNum & ".zip"
            'lDirectorio.Caption = App.Path & "\"
                
            frmMain.Inet1.Execute , "GET"
            
            Do While bDone = False
            DoEvents
            Loop
            
            If dError Then Exit Sub
            
            UnZip Directory, App.Path & "\"
            Kill Directory
        Next i
    End If
     End If
    Call GuardarInt(App.Path & "\INIT\Update.ini", iX)
    
    Image2.Enabled = True
    Call addConsole("Cliente actualizado correctamente.", 255, 255, 0, True, False)
    Call Reproducir_WAV(App.Path & "\Wav\Actualizado.wav", SND_FILENAME)
    ProgressBar1.value = 0

If MsgBox("¿Deseas Jugar?", vbYesNo) = vbYes Then
    Call ShellExecute(Me.hWnd, "open", App.Path & "/Argentum.exe", "", "", 1)
    End
 Else
    End
End If

End Sub

Private Sub Form_Load()
ProgressBar1.Picture = LoadPicture(App.Path & "\Graficos\AU_BarraVacia.jpg")
Image2.Picture = LoadPicture(App.Path & "\Graficos\AU_Buscar_N.jpg")
Image1.Picture = LoadPicture(App.Path & "\Graficos\AU_Salir_N.jpg")
frmMain.Picture = LoadPicture(App.Path & "\Graficos\AU_Main.jpg")
ProgressBar1.value = 0
'ProgressBar1.Height = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Graficos\AU_Buscar_N.jpg")
Image1.Picture = LoadPicture(App.Path & "\Graficos\AU_Salir_N.jpg")
End Sub

Private Sub Image1_Click()
Image1.Picture = LoadPicture(App.Path & "\Graficos\AU_Salir_A.jpg")
End
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Graficos\AU_Salir_A.jpg")
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Graficos\AU_Salir_I.jpg")
End Sub
Private Sub Image2_Click()
Image2.Enabled = False
Image2.Picture = LoadPicture(App.Path & "\Graficos\AU_Buscar_A.jpg")
Call Analizar

'Call addConsole("Buscando Actualizaciones...", 255, 255, 255, True, False)

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Graficos\AU_Buscar_A.jpg")
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Graficos\AU_Buscar_I.jpg")
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Select Case State
        Case icError
            Call addConsole("Error en la conexión, descarga abortada.", 255, 0, 0, True, False)
            bDone = True
            dError = True
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long
            
            FileSize = Inet1.GetHeader("Content-length")
            ProgressBar1.Max = FileSize
            
            Call addConsole("Descarga iniciada.", 0, 255, 0, True, False)
            
            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                    
                vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    ProgressBar1.value = ProgressBar1.value + Len(vtData) * 2
                    LSize.Caption = (ProgressBar1.value + Len(vtData) * 2) / 1000000 & "MBs de " & (FileSize / 1000000) & "MBs"
                    ProgressBar1.Text = "[" & CLng((ProgressBar1.value * 100) / ProgressBar1.Max) & "% Completado.]"

                    DoEvents
                Loop
            Close #1
            
            Call addConsole("Descarga finalizada", 0, 255, 0, True, False)
            LSize.Caption = FileSize & "bytes"
            ProgressBar1.value = 0
            
            bDone = True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, data
    Close #F
End Sub


