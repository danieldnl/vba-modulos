Attribute VB_Name = "mdlRef"
Option Compare Database
Option Explicit

'|--------------------------------------------------------------------------------------------------------------------|
'|----ModREF - TRATAMENTO DE BIBLIOTECAS(REFERÊNCIAS) PARA COMPATIBILIDADE ENTRE VERSÕES OFFICE E OUTRAS DEPENDÊNCIAS-|
'|--------------------------------------------------------------------------------------------------------------------|
'|--------PARA UM MELHOR FUNCIONAMENTO, DESABILITAR O FECHAMENTO DO PROGRAMA PELO BOTÃO FECHAR DO ACCESS--------------|


Function InserirReferencias()
'Aqui vão as referências complementares do programa como Selenium, PDFCreator e Office, etc, e que podem ocasionar problemas de incompatibilidade.
'Exceto a referência do próprio programa(se o programa for access por exemplo, a referência access não é necessária aqui) e a referência VBA.
'Esta função deverá ser utilizada no evento 'open' de um formulário ou planilha(não testado).

    RefPDFCreator 1
    RefSelenium 1
    RefExcel 1
    RefOutlook 1
    RefOffice 1
    RefWord 1

End Function

Function RemoverReferencias()
'Remoção das referências complementares
'Deverá ser utilizada em botões de fechamento do programa, no evento 'close' do formulário de menu

    RefPDFCreator 0
    RefSelenium 0
    RefExcel 0
    RefOutlook 0
    RefOffice 0
    RefWord 0

End Function

Function RemoverReferenciasDev()
'Usado apenas em ambiente de desenvolvimento. Para garantir que o projeto esteja compilado.
    
    If Application.IsCompiled Then
        RefPDFCreator 0
        RefSelenium 0
        RefExcel 0
        RefOutlook 0
        RefOffice 0
        RefWord 0
        MsgBox "OK"
    Else
        MsgBox "Erro. Projeto não compilado."
    End If

End Function

Function RefOffice(Valor As Integer) As Boolean

    Dim ref As Reference

    If Valor = 1 Then
        #If VB7 Then
            RefOffice = getRef("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 7)  '2013
        #Else
            RefOffice = getRef("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 3)  '2003
        #End If
    ElseIf Valor = 0 Then
        RefOffice = deleteRef("Office")
    End If
     
End Function

Function RefWord(Valor As Integer) As Boolean

    If Valor = 1 Then
        #If VB7 Then
            RefWord = getRef("{00020905-0000-0000-C000-000000000046}", 8, 6)  '2013
        #Else
            RefWord = getRef("{00020905-0000-0000-C000-000000000046}", 8, 3)  '2003
        #End If
    ElseIf Valor = 0 Then
        RefWord = deleteRef("Word")
    End If
    
End Function

Function RefExcel(Valor As Integer) As Boolean

    If Valor = 1 Then
        #If VB7 Then
            RefExcel = getRef("{00020813-0000-0000-C000-000000000046}", 1, 8)  'Excel 2013
        #Else
            RefExcel = getRef("{00020813-0000-0000-C000-000000000046}", 1, 5)  'Excel 2003
        #End If
    ElseIf Valor = 0 Then
        RefExcel = deleteRef("Excel")
    End If
    
End Function

Function RefOutlook(Valor As Integer) As Boolean

    If Valor = 1 Then
        #If VB7 Then
            RefOutlook = getRef("{00062FFF-0000-0000-C000-000000000046}", 9, 5)  '2013
        #Else
            RefOutlook = getRef("{00062FFF-0000-0000-C000-000000000046}", 9, 2)  '2003
        #End If
    ElseIf Valor = 0 Then
        RefOutlook = deleteRef("Outlook")
    End If
    
End Function

Function RefPDFCreator(Valor As Integer) As Boolean
    
    If Valor = 1 Then
        'Procura pelas versões x86 e x64, respectivamente, do PDF Creator
        If getRef("{E33847EB-154E-4A05-8222-FF4E3FD46075}", 1, 0) Or getRef("{1CE9DC08-9FBC-45C6-8A7C-4FE1E208A613}", 6, 1) Then
            RefPDFCreator = True
        Else
            RefPDFCreator = False
        End If
    ElseIf Valor = 0 Then
        RefPDFCreator = deleteRef("PDFCreator")
    End If

End Function

Function RefADO(Valor As Integer) As Boolean
    
    If Valor = 1 Then
        RefADO = getRef("{00000201-0000-0010-8000-00AA006D2EA4}", 2, 1)
    ElseIf Valor = 0 Then
        RefADO = deleteRef("ADODB")
    End If
    
End Function

Function RefOLE(Valor As Integer) As Boolean

    If Valor = 1 Then
        RefOLE = getRef("{00020430-0000-0000-C000-000000000046}", 2, 0)
    ElseIf Valor = 0 Then
        RefOLE = deleteRef("stdole")
    End If
        
End Function

Function RefDAO(Valor As Integer) As Boolean

    If Valor = 1 Then
        RefDAO = getRef("{00025E01-0000-0000-C000-000000000046}", 5, 0)
    ElseIf Valor = 0 Then
        RefDAO = deleteRef("Word")
    End If
    
End Function

Function RefExtensibility(Valor As Integer) As Boolean
    
    If Valor = 1 Then
        RefExtensibility = getRef("{0002E157-0000-0000-C000-000000000046}", 5, 3)
    ElseIf Valor = 0 Then
        RefExtensibility = deleteRef("VBIDE")
    End If
    
End Function

Function RefHtmlObject(Valor As Integer) As Boolean

    If Valor = 1 Then
        RefHtmlObject = getRef("{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}", 4, 0)
    ElseIf Valor = 0 Then
        RefHtmlObject = deleteRef("MSHTML")
    End If
    
End Function

Function RefScripting(Valor As Integer) As Boolean

    If Valor = 1 Then
        RefScripting = getRef("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
    ElseIf Valor = 0 Then
        RefScripting = deleteRef("Scripting")
    End If
    
End Function

Function RefAccess(Valor As Integer) As Boolean

    If Valor = 1 Then
        RefAccess = getRef("{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", 9, 0)
    ElseIf Valor = 0 Then
        RefAccess = deleteRef("Access")
    End If
    
End Function

Function RefInternetControls(Valor As Integer) As Boolean

    If Valor = 1 Then
        RefInternetControls = getRef("{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}", 1, 1)
    ElseIf Valor = 0 Then
        RefInternetControls = deleteRef("SHDocVw")
    End If
    
End Function

Function RefScriptHost(Valor As Integer) As Boolean

    If Valor = 1 Then
        RefScriptHost = getRef("{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}", 1, 0)
    ElseIf Valor = 0 Then
        RefScriptHost = deleteRef("IWshRuntimeLibrary")
    End If

End Function

Function RefExtra(Valor As Integer)

    If Valor = 1 Then
        RefExtra = getRef("{7FAE9440-C040-11CD-B010-0000C06E6B8A}", 7, 0)
    ElseIf Valor = 0 Then
        RefExtra = deleteRef("EXTRA")
    End If
    
End Function

Function RefSelenium(Valor As Integer)

    If Valor = 1 Then
        RefSelenium = getRef("{E57E03DE-C7FE-4C12-85C8-EC8B32DFFB86}", 2, 0)
    ElseIf Valor = 0 Then
        RefSelenium = deleteRef("SeleniumWrapper")
    End If
    
End Function

Private Function getRef(ref As String, major As Integer, minor As Integer) As Boolean

    On Error Resume Next
    References.AddFromGuid ref, major, minor
    If Err.Number = 0 Or existeRef(ref) Then
        getRef = True
    Else
        getRef = False
    End If

End Function

Private Function deleteRef(refName As String) As Boolean

    Dim vbProj As VBProject
    Dim ref As Variant
    Dim achou As Boolean
  
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    
    For Each ref In vbProj.References
        If ref.Name = refName Then
            vbProj.References.Remove (ref)
            DoEvents
            achou = True
        End If
    Next
    
    If Err.Number = 0 And achou Then
        deleteRef = True
    Else
        deleteRef = False
    End If

End Function

Public Function existeRef(refGUID As String) As Boolean

    Dim vbProj As VBProject
    Dim ref As Variant
    Dim achou As Boolean
  
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    
    For Each ref In vbProj.References
        If ref.Guid = refGUID Then
            achou = True
            Exit For
        End If
    Next
    
    existeRef = achou

End Function
