Attribute VB_Name = "mdlRequisicaoPessoas"
Option Explicit
Sub Obrigatorio()
    If _
        Sheets("REQUISIÇÃO DE PESSOAL").Range("CCATUAL").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("CARGORP").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("W3").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("K7").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("N7").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("Q7").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("K11").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("C16").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("J16").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("R16").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("C17").Value = "" Or _
        Sheets("REQUISIÇÃO DE PESSOAL").Range("C19").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("J19").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("N19").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("R19").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("O44").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("S44").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("C47").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("E47").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("G47").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("K47").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("O47").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("S47").Value = "" Or _
        Sheets("REQUISIÇÃO DE PESSOAL").Range("T74").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V74").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T75").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V75").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T76").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V76").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T77").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V77").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T78").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V78").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T79").Value = "" Or _
        Sheets("REQUISIÇÃO DE PESSOAL").Range("V79").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T80").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V80").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T81").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V81").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T82").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V82").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T83").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V83").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T84").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V84").Value = "" Or _
        Sheets("REQUISIÇÃO DE PESSOAL").Range("T85").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V85").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T86").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V86").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T87").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V87").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T88").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V88").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T89").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V89").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T90").Value = "" Or _
        Sheets("REQUISIÇÃO DE PESSOAL").Range("V90").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("T91").Value = "" Or Sheets("REQUISIÇÃO DE PESSOAL").Range("V91").Value = "" _
    Then
        MsgBox "Obrigatório o preenchimento de todos os campos em Vermelho"
        Exit Sub
    Else
        Call GestorMandaRP
    End If
End Sub
Sub GestorMandaRP()
    Dim Status As Boolean

    Status = False
    Application.ScreenUpdating = False
    Call GetUserName_Gestor
    If Sheets("REQUISIÇÃO DE PESSOAL").Range("W3").Value = "Não" Then
        Status = Gestor_BP
    ElseIf Sheets("REQUISIÇÃO DE PESSOAL").Range("W3").Value = "Sim" Then
        Status = Gestor_BP_Confidencial
    End If
    If Status Then
        Call MessageOK
        Call CloseCurrent
    End If
    Application.ScreenUpdating = True
End Sub
Sub Call_RP_BP()
    Dim strPassTry As String
    Dim strPassword As String
    Dim lTries As Long
    Dim bSuccess As Boolean
     
    Application.ScreenUpdating = False
    strPassword = strPasswordApproval
    For lTries = 1 To 3
        strPassTry = InputBox("Insira a Senha", "BP: Assinar & Enviar")
        If strPassTry = vbNullString Then Exit Sub
        bSuccess = strPassword = strPassTry
        If bSuccess = True Then Exit For
        MsgBox "Senha Incorreta"
    Next lTries
    If bSuccess = True Then Call BPMandaRP
    Application.ScreenUpdating = True
End Sub
Sub BPMandaRP()
    Dim Status As Boolean
    
    Status = False
    Call GetUserName_BP
    If Sheets("REQUISIÇÃO DE PESSOAL").Range("W3").Value = "Não" Then
        Status = BP_Facilities_IT
        If Status Then
            Status = BP_RS
        End If
    ElseIf Sheets("REQUISIÇÃO DE PESSOAL").Range("W3").Value = "Sim" Then
        Status = BP_Facilities_IT_Confidencial
        If Status Then
            Status = BP_RS_Confidencial
        End If
    End If
    If Status Then
        Call MessageOK
        Call CloseCurrent
    End If
End Sub
Sub Call_RS()
    Dim strPassTry As String
    Dim strPassword As String
    Dim lTries As Long
    Dim bSuccess As Boolean
    
    Application.ScreenUpdating = False
    strPassword = strPasswordApproval
    For lTries = 1 To 3
        strPassTry = InputBox("Insira a Senha", "BP: Assinar & Enviar")
        If strPassTry = vbNullString Then Exit Sub
        bSuccess = strPassword = strPassTry
        If bSuccess = True Then Exit For
        MsgBox "Senha Incorreta"
    Next lTries
    If bSuccess = True Then Call RSMandaFolha
    Application.ScreenUpdating = True
End Sub
Sub RSMandaFolha()
    Dim Status As Boolean
    
    Status = False
    If Sheets("REQUISIÇÃO DE PESSOAL").Range("W3").Value = "Não" Then
        Status = RS_Folha
        If Status Then
            Status = RS_Informa
        End If
    ElseIf Sheets("REQUISIÇÃO DE PESSOAL").Range("W3").Value = "Sim" Then
        Status = RS_Folha_Confidencial
        If Status Then
            Status = RS_Informa_Confidencial
        End If
    End If
    If Status Then
        Call foo
        Call foo2
        Call MessageOK
        Call CloseCurrent
    End If
End Sub
Function Gestor_BP() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    Gestor_BP = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("REQUISIÇÃO DE PESSOAL").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Requisição de Pessoal: " & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7")
    MailTxt = "Olá!," & vbNewLine & vbNewLine & "Segue em anexo a Requisição de Pessoal para " & Range("CARGORP") & " " & "em" & " " & Range("N7") & "/" & Range("Q7") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP = SendEmailRP("RP_GESTOR", MailTo, MailSub, MailTxt)
End Function
Function Gestor_BP_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    Gestor_BP_Confidencial = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("REQUISIÇÃO DE PESSOAL").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Requisição de Pessoal Confidencial"
    MailTxt = "Olá!," & vbNewLine & vbNewLine & "Segue em anexo Requisição de Pessoal confidencial para " & Range("CARGORP") & " " & "em" & " " & Range("N7") & "/" & Range("Q7") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP_Confidencial = SendEmailRP("RP_GESTOR_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    BP_Facilities_IT = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailRP_BP_Facilities_IT, "RP")
    MailSub = "JML - Requisição de Materiais & Acessos: " & Range("X11") & "_" & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue abertura de chamado de Materiais & Acessos para o formulário " & Range("X11") & ", " & Range("CARGORP") & " " & "de" & " " & Range("N7") & " " & Range("Q7") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT = SendEmailRP("RP_BP", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    BP_Facilities_IT_Confidencial = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailRP_BP_Facilities_IT, "RP")
    MailSub = "JML - Requisição de Materiais & Acessos Confidencial: " & Range("X11")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue abertura de chamado de Materiais & Acessos confidencial para o formulário " & Range("X11") & ", " & Range("CARGORP") & " " & "de" & " " & Range("N7") & " " & Range("Q7") & "." & " A vaga correspondente está sendo aberta em caráter de confidencialidade e deve ser tratada com discrição." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT_Confidencial = SendEmailRP("RP_BP_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function BP_RS() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    BP_RS = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = strEmailRP_BP_RS
    MailSub = "JML - Requisição de Pessoal: " & Range("X11") & " " & Range("CARGORP") & " " & Range("N7") & " " & Range("Q7")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue em anexo o formulário " & Range("X11") & ", Requisição de Pessoal para" & " " & Range("CARGORP") & " " & "em" & " " & Range("N7") & "-" & Range("Q7") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_RS = SendEmailRP("RP_BP_RS", MailTo, MailSub, MailTxt)
End Function
Function BP_RS_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    BP_RS_Confidencial = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = strEmailRP_BP_RS
    MailSub = "JML - Requisição de Pessoal Confidencial: " & Range("X11")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue em anexo o formulário " & Range("X11") & ", Requisição de Pessoal confidencial para" & " " & Range("CARGORP") & " " & "em" & " " & Range("N7") & "-" & Range("Q7") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_RS_Confidencial = SendEmailRP("RP_BP_RS_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = strEmailRP_RS_Folha
    MailSub = "JML - Requisição de Pessoal: " & Range("X11") & "_" & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue Requisição de Pessoal para " & Range("CARGORP") & " " & "de" & " " & Range("N7") & "-" & Range("Q7") & "." & " O formulário " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Folha = SendEmailRP("RP_RS", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha_Confidencial = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = strEmailRP_RS_Folha
    MailSub = "JML - Requisição de Pessoal Confidencial: " & Range("X11")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue Requisição de Pessoal confidencial para " & Range("CARGORP") & " " & "de" & " " & Range("N7") & "-" & Range("Q7") & "." & " O formulário " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Folha_Confidencial = SendEmailRP("RP_RS_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function RS_Informa() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Informa = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailRP_RS_Informa, "RP")
    MailSub = "JML - Finalização do formulário " & Range("X11") & ": " & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Informamos que a vaga de " & Range("CARGORP") & " " & Range("N7") & "/" & Range("Q7") & " formulário " & Range("X11") & ", foi preenchida com o(a) candidato(a) " & Range("L32") & " que iniciará em " & Range("W32") & "." & vbNewLine & "Atual funcionário companyname? " & Range("F32") & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Informa = SendEmailRP("RP_RS_INFORMA", MailTo, MailSub, MailTxt)
End Function
Function RS_Informa_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Informa_Confidencial = False
    ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailRP_RS_Informa, "RP")
    MailSub = "JML - Finalização do formulário confidencial " & Range("X11")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Informamos que a vaga confidencial de " & Range("CARGORP") & " " & Range("N7") & "/" & Range("Q7") & " formulário " & Range("X11") & ", foi preenchida com o(a) candidato(a) " & Range("L32") & " que iniciará em " & Range("W32") & "." & vbNewLine & "Atual funcionário companyname? " & Range("F32") & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Informa_Confidencial = SendEmailRP("RP_RS_INFORMA_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Sub foo()
    Dim x As Workbook
    Dim y As Workbook
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set x = ActiveWorkbook
    If strAmbiente = "producao" Then
        Set y = Workbooks.Open("\\sbra080155\rh$\BASE JML\Base Joiners.xlsx")
    Else
        Set y = Workbooks.Open("\\sbra080155\public$\BI & Systems\BASE JML\Base Joiners.xlsx")
    End If
    x.Sheets("RP").Range("B4:EE4").Copy
    Range("B" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats
    Selection.PasteSpecial Paste:=xlPasteColumnWidths
    y.Save
    y.Close
End Sub
Sub foo2()
    Dim x As Workbook
    Dim y As Workbook
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set x = ActiveWorkbook
    
    If strAmbiente = "producao" Then
        Set y = Workbooks.Open("\\sbra080155\rh$\BASE JML - Facilities\Base Joiners - Facilities.xlsx")
    Else
        Set y = Workbooks.Open("\\sbra080155\public$\BI & Systems\BASE JML - Facilities\Base Joiners - Facilities.xlsx")
    End If
    x.Sheets("RP Facilities").Range("B4:AB4").Copy
    Range("B" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats
    Selection.PasteSpecial Paste:=xlPasteColumnWidths
    y.Save
    y.Close
End Sub
Function SendEmailRP(pType As String, pTo As String, pSubject As String, pBody As String) As Boolean
    Dim wb1 As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim wsI As Worksheet
    Dim wsO As Worksheet
    Dim TempFileNameConfidencial As String
    
    'Os tipos de forma de envio de email (pType) são:
    'RP 1º PASSO) RP_GESTOR e RP_GESTOR_CONFIDENCIAL
    'RP 2º PASSO) RP_BP e RP_BP_CONFIDENCIAL
    'RP 2º PASSO) RP_BP_RS e RP_BP_RS_CONFIDENCIAL
    'RP 3º PASSO) RP_RS e RP_RS_CONFIDENCIAL
    'RP 3º PASSO) RP_RS_INFORMA e RP_RS_INFORMA_CONFIDENCIAL
    SendEmailRP = False
    On Error GoTo TrataErro
    With Application
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    ActiveWorkbook.Unprotect Password:=strPasswordLock
    ActiveSheet.Unprotect Password:=strPasswordLock
    TempFilePath = Environ$("temp") & "\"
    FileExtStr = "." & LCase(Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".", , 1)))
    If InStr(1, pType, " CONFIDENCIAL") > 0 Then
        TempFileNameConfidencial = " Confidencial"
    ElseIf pType = "RP_BP_CONFIDENCIAL" Then
        TempFileNameConfidencial = ""
    End If
    
    If pType = "RP_GESTOR" Or pType = "RP_GESTOR_CONFIDENCIAL" Then
        TempFileName = "JML - Requisição de Pessoal" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = ActiveWorkbook
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlMovimentacao")
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlSolicitacaoDesligamento")
        wb1.Sheets(Array("SOLICITAÇÃO DE DESLIGAMENTO", "MOVIMENTAÇÃO", "SD", "MOV", "SD Facilities", "Mov Facilities")).Delete
        wb1.SaveCopyAs TempFilePath & TempFileName & FileExtStr
        On Error GoTo TrataErro
    ElseIf pType = "RP_BP" Or pType = "RP_BP_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Requisição de Materias & Acessos" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL")
        With wb1
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayWorkbookTabs = False
            ActiveWindow.DisplayHeadings = False
            ActiveWindow.Zoom = 90
            Set wsO = wb1.ActiveSheet
            wsI.Rows("34:96").EntireRow.Copy
            wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            wsO.Range("A2").PasteSpecial Paste:=xlPasteColumnWidths
            wsO.Range("A2").PasteSpecial Paste:=xlPasteFormats
            wsO.Range("A2").Select
        End With
        wb1.SaveAs Filename:=TempFilePath & TempFileName & FileExtStr, FileFormat:=xlOpenXMLWorkbook
        wb1.Close
        On Error GoTo TrataErro
    ElseIf pType = "RP_BP_RS" Or pType = "RP_BP_RS_CONFIDENCIAL" Then
        TempFileName = "JML - Requisição de Pessoal" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = ActiveWorkbook
        wb1.SaveCopyAs TempFilePath & TempFileName & FileExtStr
        On Error GoTo TrataErro
    ElseIf pType = "RP_RS" Or pType = "RP_RS_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Requisição de Pessoal" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL")
        With wb1
            ActiveWindow.DisplayWorkbookTabs = False
            ActiveWindow.DisplayHeadings = False
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 90
            Set wsO = wb1.ActiveSheet
            wsI.Rows("2:33").EntireRow.Copy
            wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            wsO.Range("A2").PasteSpecial Paste:=xlPasteColumnWidths
            wsO.Range("A2").PasteSpecial Paste:=xlPasteFormats
            wsI.Shapes("Imagem 1").Copy
            wsO.Range("B2").Select
            wsO.Paste
        End With
        wb1.SaveAs Filename:=TempFilePath & TempFileName & FileExtStr, FileFormat:=xlOpenXMLWorkbook
        wb1.Close
        On Error GoTo TrataErro
    ElseIf pType = "RP_RS_INFORMA" Or pType = "RP_RS_INFORMA_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Requisição de Materias & Acessos" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("x11") & "_" & Range("CARGORP") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL")
        With wb1
           ActiveWindow.DisplayGridlines = False
           ActiveWindow.DisplayWorkbookTabs = False
           ActiveWindow.DisplayHeadings = False
           ActiveWindow.Zoom = 90
           Set wsO = wb1.ActiveSheet
           wsI.Rows("34:96").EntireRow.Copy
           wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
           SkipBlanks:=False, Transpose:=False
           wsO.Range("A2").PasteSpecial Paste:=xlPasteColumnWidths
           wsO.Range("A2").PasteSpecial Paste:=xlPasteFormats
           wsO.Range("A2").Select
        End With
        wb1.SaveAs Filename:=TempFilePath & TempFileName & FileExtStr, FileFormat:=xlOpenXMLWorkbook
        wb1.Close
        On Error GoTo TrataErro
    End If
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect Password:=strPasswordLock, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveWorkbook.Protect Password:=strPasswordLock, Structure:=True, Windows:=False
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = pTo
        .CC = ""
        .BCC = ""
        .Subject = pSubject
        .Body = pBody
        .Attachments.Add TempFilePath & TempFileName & FileExtStr
        .Send  'or use .Send
    End With
    Kill TempFilePath & TempFileName & FileExtStr
    Set OutMail = Nothing
    Set OutApp = Nothing
    With Application
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    SendEmailRP = True
    Exit Function
TrataErro:
    With Application
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect Password:=strPasswordLock, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveWorkbook.Protect Password:=strPasswordLock, Structure:=True, Windows:=False
    MsgBox "Ocorreu um erro ao enviar o email. Por favor, contate o administrador. Descrição : " & Err.Description
End Function
