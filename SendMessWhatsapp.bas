' Require Variants to be declared before used
Option Explicit

Public WrkB As Workbook
Public WrkS As Worksheet

Public IntervaloRotina As Range
Public Celula          As Range
Public oShell As Object
Public oAutoIt As Object
Public script As Object

Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
Public Sub EnvioWhatsapp()



Set WrkB = ThisWorkbook
Set WrkS = WrkB.Sheets("Whatsapp")

Set IntervaloRotina = WrkS.Range("A3:A100000")


Set oShell = CreateObject("Wscript.shell")
Set oAutoIt = CreateObject("AutoItX3.Control")
oShell.Run "chrome.exe --start-maximized", 2, False

Sleep (2000)
oAutoIt.WinWaitActive "Nova guia - Google Chrome" 'Aguardar o chrome abrir
Sleep (10000)
oAutoIt.Send "https://web.whatsapp.com/{ENTER}"
oAutoIt.Sleep 10000 'Aguardar escanear código

With WrkS
    .Select
        For Each Celula In IntervaloRotina
             
        
            Call Rotina
            Sleep (5000)
            Next
        
End With

End Sub
Public Sub Rotina()

If WrkS.Cells(Celula.Row, 1).Value = "" Then
   MsgBox "Procedimento Finalizado"
   End
End If

Dim telefone As String
Dim mensagem As String

Dim endereco As String

telefone = WrkS.Cells(Celula.Row, 1).Value
mensagem = WrkS.Cells(Celula.Row, 2).Value


endereco = "https://web.whatsapp.com/send?phone=" & telefone & "&text=" & mensagem

Sleep (7000)
oAutoIt.Send "!d"
Sleep (1000)
oAutoIt.Send endereco
Sleep (1000)
oAutoIt.Send "{ENTER}"
Sleep (10000)
oAutoIt.Send "{ENTER}"



'oAutoIt.WinWaitActive "Bloco de Notas"
'oAutoIt.Send "!n"
'oAutoIt.WinWaitClose "Sem título - Bloco de Notas", "", 10

'oShell.Quit
'oShell = Nothing
'oAutoIt = Nothing

WrkS.Cells(Celula.Row, 3) = "Enviado com Sucesso"

End Sub
