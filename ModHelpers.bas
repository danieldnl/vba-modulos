Attribute VB_Name = "modHelpers"
Option Compare Database

#If Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Function ApenasNumeros(strPalavra) As String

    Dim tamanho As Integer
    Dim i As Integer
    Dim numeros As String
    
    tamanho = Len(strPalavra)
    strPalavra = UCase(strPalavra)
    If InStr(1, strPalavra, "ÚNICA", vbTextCompare) Then 'É NECESSÁRIO PARA RECONHECER A VARA ÚNICA
        numeros = 1
    Else 'PEGA NO TEXTO O VALOR DA VARA
        For i = 1 To tamanho
            If IsNumeric(Mid(strPalavra, i, 1)) = True Then
                numeros = numeros & Mid(strPalavra, i, 1)
            End If
        Next
    End If
    ApenasNumeros = numeros
    
End Function

Public Function ConverterSegParaHMS(Segundos As Long)

    ConverterSegParaHMS = Format(Segundos / 24 / 60 / 60, "hh:mm:ss")

End Function

Public Function DigVerifCNPJ(cnpj As String) As String
'Calcula os dígitos verificadores do CNPJ
    
    Dim i As Integer
    Dim intFator As Integer
    Dim intTotal As Integer
    Dim intResto
        
    'Verifica se tem 12 ou 14 dígitos
    If Not (Len(cnpj) = 12 Or Len(cnpj) = 14) Then
        Exit Function
    Else
        'Verifica se é numérico
        If Not IsNumeric(cnpj) Then
            Exit Function
        Else
            'Trunca o CNPJ em 12 caracteres
            cnpj = Left$(cnpj, 12)
        End If
    End If

Inicio:
    'Percorre as colunas (de trás para frente),
    'multiplicando por seus respectivos fatores
    intFator = 2
    intTotal = 0
    For i = Len(cnpj) To 1 Step -1
        If intFator > 9 Then intFator = 2
        intTotal = intTotal + ((CInt(Mid(cnpj, i, 1)) * intFator))
        intFator = intFator + 1
    Next i

    'Obtém o resto da divisão por 11
    i = intTotal Mod 11
    'Subtrai 11 do resto
    i = 11 - i
    'O dígito verificador é i
    If i = 10 Or i = 11 Then i = 0
    'Concatena ao CNPJ
    cnpj = cnpj & CStr(i)
    If Len(cnpj) = 13 Then
        'Calcula o segundo dígito
        GoTo Inicio
    End If
    'Retorna os dígitos verificadores
    DigVerifCNPJ = Right$(cnpj, 2)
    
End Function

Public Sub FecharBrowserOculto(browser As String)
    
    Dim ws As New WshShell
    
    'Variáveis para browser:
    'iexplore.exe
    'chrome.exe
    'firefox.exe
    
    On Error GoTo Err
    ws.Run "cmd /C TASKKILL /IM " & browser & " /F", 0, True

Err:
    If Err.Number <> 0 Then
        MsgBox "Erro nº " & Err.Number & vbCrLf & Err.Description, vbCritical, "Erro"
    End If
    
End Sub

Public Sub FecharConexaoADO(ByRef con As ADODB.Connection)

    If Not con Is Nothing Then
        If con.State = 1 Then
            con.Close
        End If
        Set con = Nothing
    End If

End Sub

Public Function makeObject(stObj As String) As Object
'Singleton
    
    Dim obj As Object
    
    On Error Resume Next
    
    Set obj = GetObject(, stObj)
    If Err.Number = 429 Then
        Set obj = CreateObject(stObj)
    End If
    
    Set makeObject = obj
    
End Function

Public Function RemoveCaracteres(stChar As String) As String

    Dim chrEspeciais, chrBasicos As String
    Dim i As Integer, p As String
 
    'Acentos e caracteres especiais que serão buscados na string
    'Você pode definir outros caracteres nessa variável, mas
    ' precisará também colocar a letra correspondente em chrBasicos
    'Vazios serão tirados ao final
    
    chrEspeciais = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
     
    'Letras correspondentes para substituição
    chrBasicos = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
     
    'Loop que irá de andará a string letra a letra
    For i = 1 To Len(stChar)
     
        'InStr buscará se a letra indice i de stChar pertence a
        ' chrEspeciais e se existir retornará a posição dela
        p = InStr(chrEspeciais, Mid(stChar, i, 1))
         
        'Substitui a letra de indice i em chrEspeciais pela sua
        ' correspondente em chrBasicos
        If p > 0 Then
            Mid(stChar, i, 1) = Mid(chrBasicos, p, 1)
        End If
    Next
     
    'Retorna a nova string
    RemoveCaracteres = stChar
    
End Function

Public Function RemoveCaracteresExt(st As String) As String
    
    Dim stChar As String
    
    stChar = RemoveCaracteres(st)
    stChar = SomenteLetras(stChar)
    stChar = Trim(stChar)
    stChar = Replace(stChar, "   ", " ")
    stChar = Replace(stChar, "  ", " ")
    RemoveCaracteresExt = stChar
    
End Function

Public Sub SelecionarPasta(ctl As control) '-> ATENÇÃO: ctl SÓ PODE SER TextBox OU ComboBox

    Dim stPst As String, pstDlg As Office.FileDialog
    stPst = CurDir & "\"
    Set pstDlg = Application.FileDialog(msoFileDialogFolderPicker)
    With pstDlg
        .AllowMultiSelect = False
        .Title = "Selecione o arquivo"
        .InitialFileName = stPst
        .Filters.Clear
        If .Show Then
            ctl.SetFocus
            ctl.Text = .SelectedItems(1)
        End If
    End With
    
End Sub

Public Function SelecionarPastaComRetorno() As String
    Dim dlgOpen As FileDialog
    Dim escolha As MsoAlertCancelType
    
    Set dlgOpen = Application.FileDialog(msoFileDialogFolderPicker)
    With dlgOpen
        If .Show = -1 Then
            .AllowMultiSelect = False
            .Title = "Selecionar pasta para salvar os arquivos"
            SelecionarPastaComRetorno = .SelectedItems(1)
        Else
            End
        End If
    End With

    SelecionarPastaComRetorno = dlgOpen.SelectedItems(1) & "\" 'RETORNA O CAMINHO COMPLETO DO ARQUIVO, INCLUSIVE COM O NOME E EXTENSÃO

End Function

Function SomenteAlgarismos(strTexto) As String

    Dim x      As Integer
    Dim strChar As String * 1
    For x = 1 To Len(strTexto)
        strChar = Mid(strTexto, x, 1)
        If strChar Like "[0-9]" Then
            SomenteAlgarismos = SomenteAlgarismos & strChar
        End If
    Next
    
End Function

Public Function SomenteLetras(st As String) As String

    Dim i As Integer, stAux As String * 1
    For i = 1 To Len(st)
        stAux = Mid(st, i, 1)
        If stAux Like "[A-Z]" Or stAux = " " Then SomenteLetras = SomenteLetras & stAux
    Next
    
End Function
