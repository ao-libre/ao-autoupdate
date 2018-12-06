VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "VBALPROGBAR6.OCX"
Begin VB.Form frmLauncher 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet InetGithubReleases 
      Left            =   1080
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetGithubAutoupdate 
      Left            =   360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBoxLog 
      Height          =   2775
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   64
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmLauncher.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton BtnWorldEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      Caption         =   "WORLDEDITOR"
      Height          =   975
      Left            =   9120
      TabIndex        =   2
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton BtnServer 
      Caption         =   "SERVER"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton BtnGame 
      Caption         =   "JUEGO"
      Height          =   975
      Left            =   5640
      TabIndex        =   0
      Top             =   6240
      Width           =   1815
   End
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar1 
      Height          =   450
      Left            =   3000
      TabIndex        =   5
      Top             =   5040
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   794
      Picture         =   "frmLauncher.frx":007F
      BackColor       =   4194368
      ForeColor       =   0
      BorderStyle     =   0
      BarPicture      =   "frmLauncher.frx":009B
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
   Begin VB.Label LblSpanish 
      Caption         =   "Spanish"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label LblEnglish 
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label LblVersion 
      Caption         =   "Version"
      Height          =   375
      Left            =   12480
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Dim Directory As String, bDone As Boolean, dError As Boolean, F As Integer

Dim SizeInMb As Double
Dim JsonObject As Object

Private Language As String
Private JsonLanguage As Object

Private Sub BtnGame_Click()
    Call Analizar("Client")
End Sub

Private Sub BtnServer_Click()
    Call Analizar("Server")
End Sub

Private Sub BtnWorldeditor_Click()
    Call Analizar("Worldeditor")
End Sub

Private Sub LblEnglish_Click()
    Call WriteVar(App.Path & "\ConfigAutoupdate.ini", "ConfigAutoupdate", "language", "english")
    Call LaunchPopUpBeforeClose
End Sub

Private Sub LblSpanish_Click()
    Call WriteVar(App.Path & "\ConfigAutoupdate.ini", "ConfigAutoupdate", "language", "spanish")
    Call LaunchPopUpBeforeClose
End Sub

Private Sub LaunchPopUpBeforeClose()
    If MsgBox(JsonLanguage.Item("close_before_change_language"), vbYesNo) = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    LblVersion.Caption = "v" & App.Major & "." & App.Minor
    Call SetLanguageApplication
    Call CheckIfIEVersionIsCompatible
    Call CheckIfRunningLastVersionAutoupdate

    BtnGame.Caption = JsonLanguage.Item("game")
    BtnServer.Caption = JsonLanguage.Item("server")
    BtnWorldEditor.Caption = JsonLanguage.Item("worldeditor")
    LblSpanish.Caption = JsonLanguage.Item("spanish_label")
    LblEnglish.Caption = JsonLanguage.Item("english_label")
    ProgressBar1.Text = JsonLanguage.Item("completed")
End Sub

Private Sub SetLanguageApplication(Optional LanguageSelection As String)
    Dim JsonLanguageString As String
    
    If LanguageSelection = "" Then
        Language = GetVar(App.Path & "\ConfigAutoupdate.ini", "ConfigAutoupdate", "language")
    Else
        Language = LanguageSelection
    End If
    
    JsonLanguageString = FileToString(App.Path & "\Languages\" & Language & ".json")
    Set JsonLanguage = JSON.parse(JsonLanguageString)
End Sub

Public Function GetIEVersion()
    Dim FileSystemObject As New FileSystemObject
    Dim Version As String
    
    Version = FileSystemObject.GetFileVersion("c:\windows\system32\ieframe.dll")
    GetIEVersion = Version
End Function

Public Function CheckIfIEVersionIsCompatible()
    Dim IEVersion As String
    Dim IEVersionArray() As String

    IEVersion = GetIEVersion
    IEVersionArray() = Split(IEVersion, ".")

    If CInt(IEVersionArray(0)) < 10 Then
        Dim windowsXpTutorial As String
        windowsXpTutorial = GetVar(App.Path & "\ConfigAutoupdate.ini", "Links", "windowsXpTutorial")
        MsgBox (Replace(JsonLanguage.Item("error_ie"), "VAR_IEVersion", IEVersionArray(0)))
        MsgBox (JsonLanguage.Item("error_windows_xp") & windowsXpTutorial)
        End
    End If
End Function

Private Function FileToString(strFilename As String) As String
    IFile = FreeFile
    Open strFilename For Input As #IFile
        FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
End Function

Private Sub CheckIfRunningLastVersionAutoupdate()
    Dim responseGithub As String, versionNumberMaster As String, versionNumberLocal As String
    Dim githubAccount As String
    
    githubAccount = GetVar(App.Path & "\ConfigAutoupdate.ini", "ConfigAutoupdate", "githubAccount")

    responseGithub = InetGithubAutoupdate.OpenURL("https://api.github.com/repos/" & githubAccount & "/ao-autoupdate/releases/latest")
    Set JsonObject = JSON.parse(responseGithub)
    
    versionNumberMaster = JsonObject.Item("tag_name")
    versionNumberLocal = GetVar(App.Path & "\ConfigAutoupdate.ini", "ConfigAutoupdate", "version")
    
    If Not versionNumberMaster = versionNumberLocal Then
        MsgBox (JsonLanguage.Item("launcher_outdated"))
        MsgBox (JsonLanguage.Item("your_version") & " " & versionNumberLocal & " " & JsonLanguage.Item("last_version") & " " & versionNumberMaster)
        End
    End If
End Sub

Private Function CheckIfApplicationIsUpdated(ApplicationToUpdate As String) As Boolean
    Dim versionNumberLocal As String, versionNumberMaster As String
    Dim repository As String, githubAccount As String
    Dim responseGithub As String, urlEndpointUpdate As String, fileToExecuteAfterUpdated As String
    
    githubAccount = GetVar(App.Path & "\ConfigAutoupdate.ini", ApplicationToUpdate, "githubAccount")

    Call addConsole(JsonLanguage.Item("looking_for_upgrades"), 255, 255, 255, True, False)
    Call addConsole(JsonLanguage.Item("configured_to") & ApplicationToUpdate, 100, 200, 40, True, False)   '>> Informacion
    
    Call Reproducir_WAV(App.Path & "\Wav\Revision.wav", SND_FILENAME)
    
    repository = GetVar(App.Path & "\ConfigAutoupdate.ini", ApplicationToUpdate, "repository")
    urlEndpointUpdate = "https://api.github.com/repos/" & githubAccount & "/" & repository & "/releases/latest"
    
    responseGithub = InetGithubReleases.OpenURL(urlEndpointUpdate)

    Set JsonObject = JSON.parse(responseGithub)
    versionNumberMaster = JsonObject.Item("tag_name")
    versionNumberLocal = GetVar(App.Path & "\ConfigAutoupdate.ini", ApplicationToUpdate, "version")

    If versionNumberMaster = versionNumberLocal Then
        CheckIfApplicationIsUpdated = True
    ElseIf Not versionNumberMaster = versionNumberLocal Then
        CheckIfApplicationIsUpdated = False
    End If
    
End Function

Private Function GetGithubReleaseData() As Object

End Function

Private Sub Analizar(ApplicationToUpdate As String)
    Dim SubDirectoryApp As String
    Dim isApplicationUpdated As Boolean
    
    isApplicationUpdated = CheckIfApplicationIsUpdated(ApplicationToUpdate)
    
    If isApplicationUpdated = True Then
        Call addConsole(JsonLanguage.Item("up_to_date"), 149, 100, 210, True, False)
    Else
        If MsgBox(JsonLanguage.Item("download_continue"), vbYesNo) = vbYes Then
            ProgressBar1.Visible = True
            
            Call addConsole(JsonLanguage.Item("starting"), 200, 200, 200, True, False)   '>> Informacion
            
            ProgressBar1.Max = JsonObject.Item("assets").Item(1).Item("size")
            SizeInMb = BytesToMegabytes(JsonObject.Item("assets").Item(1).Item("size"))
            
            InetGithubAutoupdate.AccessType = icUseDefault
            InetGithubAutoupdate.URL = JsonObject.Item("assets").Item(1).Item("browser_download_url")
            Directory = App.Path & "\Updates\" & JsonObject.Item("assets").Item(1).Item("name")
            bDone = False
            dError = False
                
            InetGithubAutoupdate.Execute , "GET"
            
            Do While bDone = False
                DoEvents
            Loop
                
            If dError Then Exit Sub
        
            SubDirectoryApp = GetVar(App.Path & "\ConfigAutoupdate.ini", ApplicationToUpdate, "folderToExtract")
            
            Call addConsole(JsonLanguage.Item("one_more_moment"), 50, 90, 220, True, False)
            UnZip Directory, App.Path & "\" & SubDirectoryApp
            Kill Directory
            
            Call WriteVar(App.Path & "\ConfigAutoupdate.ini", ApplicationToUpdate, "version", CStr(JsonObject.Item("tag_name")))
            
            Call addConsole(ApplicationToUpdate & JsonLanguage.Item("update_succesful"), 66, 255, 30, True, False)
            Call addConsole(JsonLanguage.Item("comments_update") & JsonObject.Item("body") & ".", 200, 200, 200, True, False)
            Call Reproducir_WAV(App.Path & "\Wav\Actualizado.wav", SND_FILENAME)
            ProgressBar1.Value = 0
            
        ElseIf vbNo Then
            Call addConsole(JsonLanguage.Item("download_canceled"), 255, 0, 0, True, False)
        End If
    End If

    If MsgBox(Replace(JsonLanguage.Item("open_app"), "VAR_Program", ApplicationToUpdate), vbYesNo) = vbYes Then
        fileToExecuteAfterUpdated = GetVar(App.Path & "\ConfigAutoupdate.ini", ApplicationToUpdate, "fileToExecuteAfterUpdated")
        
        If LenB(SubDirectoryApp) < 0 Then
        
            Call ShellExecute(Me.hWnd, "open", App.Path & "\" & SubDirectoryApp & fileToExecuteAfterUpdated, "", "", 1)
        Else
            Call ShellExecute(Me.hWnd, "open", App.Path & "\" & fileToExecuteAfterUpdated, "", "", 1)
        End
     Else
        End
    End If

End Sub

Public Function BytesToMegabytes(Bytes As Double) As Double
   'This function gives an estimate to two decimal
   'places.  For a more precise answer, format to
   'more decimal places or just return dblAns
 
  Dim dblAns As Double
  dblAns = (Bytes / 1024) / 1024
  BytesToMegabytes = Format(dblAns, "###,###,##0.00")
  
End Function


Private Sub InetGithubAutoupdate_StateChanged(ByVal State As Integer)
    Dim Percentage As Long
    Select Case State
        Case icError
            Call addConsole(JsonLanguage.Item("error_connection"), 255, 0, 0, True, False)
            bDone = True
            dError = True
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Call addConsole(JsonLanguage.Item("download_started"), 100, 255, 130, True, False)
            
            Open Directory For Binary Access Write As #1
                vtData = InetGithubAutoupdate.GetChunk(1024, icByteArray)
                DoEvents
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                    
                    vtData = InetGithubAutoupdate.GetChunk(1024, icByteArray)

                    ProgressBar1.Value = ProgressBar1.Value + Len(vtData) * 2
                    
                    Percentage = (ProgressBar1.Value / ProgressBar1.Max) * 100
                    ProgressBar1.Text = "[" & Percentage & "% de " & SizeInMb & " MBs.]"
                    
                    DoEvents
                Loop
            Close #1
            
            Call addConsole(JsonLanguage.Item("download_finished"), 0, 255, 0, True, False)

            ProgressBar1.Value = 0
            
            bDone = True
        Case icRequesting
            'Call addConsole("Buscando ultima version disponible", 0, 76, 0, True, False)
        Case icConnecting
            'Call addConsole("Obteniendo numero de la ultima actualizacion ¯\_(O.O)_/¯", 0, 255, 0, True, False)
        Case 1 'icHostResolvingHost
            'Call addConsole("Resolviendo host... por favor espere", 0, 130, 0, True, False)
        Case icRequestSent
            'Call addConsole("Seguimos resolviendo host..", 110, 230, 20, True, False)
        Case icReceivingResponse
            'Call addConsole("Escuchamos una señal, vamos a comprobar que tengas la ultima version.", 100, 190, 200, True, False)
        Case icConnected
            'Call addConsole("Nos conectamos, ya vamos a empezar a bajar... paciencia =P ", 200, 90, 220, True, False)
        Case icResponseReceived
            'Call addConsole("Recibimos respuesta", 250, 140, 10, True, False)
        Case icHostResolved
            'Call addConsole("Lo hicimos resolvimos el host.", 110, 30, 20, True, False)
        Case Else
            Dim WebpageAolibre As String
            WebpageAolibre = GetVar(App.Path & "\ConfigAutoupdate.ini", "Links", "webpage")
            Call addConsole(JsonLanguage.Item("error_connection") & WebpageAolibre, 255, 0, 0, True, False)
    End Select
End Sub
