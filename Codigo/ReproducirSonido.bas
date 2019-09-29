Attribute VB_Name = "ReproducirSonido"
' Constantes para los flags

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  look for application specific association
Private Const SND_APPLICATION = &H80

'  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS = &H10000

'  name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS_ID = &H110000

'  play asynchronously
Private Const SND_ASYNC = &H1

  '  play synchronously (default)
Private Const SND_SYNC = &H0

'  name is a file name
Public Const SND_FILENAME = &H20000

'  loop the sound until next sndPlaySound
Private Const SND_LOOP = &H8

'  lpszSoundName points to a memory file
Private Const SND_MEMORY = &H4

'  silence not default, if sound not found
Private Const SND_NODEFAULT = &H2

 '  don't stop any currently playing sound
Private Const SND_NOSTOP = &H10

 '  don't wait if the driver is busy
Private Const SND_NOWAIT = &H2000

 '  purge non-static events for task
Private Const SND_PURGE = &H40

 '  name is a resource name or atom
Private Const SND_RESOURCE = &H40004

' Declaración del api PlaySound
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

' Reproduce el archivo de sonido wav
Public Sub Reproducir_WAV(Archivo As String, Flags As Long)
    
    Dim ret As Long
    ' Le pasa el path y los flags al api
    ret = PlaySound(Archivo, ByVal 0&, Flags)
End Sub
