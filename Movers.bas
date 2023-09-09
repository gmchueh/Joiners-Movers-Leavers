Attribute VB_Name = "mdlMovimentacao"
Option Explicit
Sub Obrigatorio_MOV()
    If _
        Sheets("MOVIMENTAÇÃO").Range("CARGOMOV").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("CCATUAL").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("W3").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("J7").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("N7").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("Q7").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("J11").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("P11").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("C16").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("J16").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("Q16").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("C19").Value = "" Or _
        Sheets("MOVIMENTAÇÃO").Range("J19").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("Q19").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("C22").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("J22").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("O22").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("Q22").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("C25").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("J25").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("O25").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("Q25").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("C28").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("J28").Value = "" Or _
        Sheets("MOVIMENTAÇÃO").Range("Q28").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V28").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V57").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("C66").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("E66").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("G66").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("K66").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("U66").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("C69").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("E69").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("G69").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("K69").Value = "" Or _
        Sheets("MOVIMENTAÇÃO").Range("U69").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T96").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V96").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T97").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V97").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T98").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V98").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T99").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V99").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T100").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V100").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T101").Value = "" Or _
        Sheets("MOVIMENTAÇÃO").Range("V101").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T102").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V102").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T103").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V103").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T104").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V104").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T105").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V105").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T106").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V106").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T107").Value = "" Or _
        Sheets("MOVIMENTAÇÃO").Range("V107").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T108").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V108").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T109").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V109").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T110").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V110").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T111").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V111").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T112").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V112").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("T113").Value = "" Or Sheets("MOVIMENTAÇÃO").Range("V113").Value = "" _
    Then
        MsgBox "Obrigatório o preenchimento de todos os campos em Vermelho"
        Exit Sub
    Else
        Call GestorMandaRP_MOV
    End If
End Sub
Sub GestorMandaRP_MOV()
    Dim Status As Boolean
    
    Status = False
    Application.ScreenUpdating = False
    Call GetUserName_Gestor
    If Sheets("MOVIMENTAÇÃO").Range("W3").Value = "Não" Then
        Status = Gestor_BP_MOV
    ElseIf Sheets("MOVIMENTAÇÃO").Range("W3").Value = "Sim" Then
        Status = Gestor_BP_MOV_Confidencial
    End If
    If Status Then
        Call MessageOK
        Call CloseCurrent
    End If
    Application.ScreenUpdating = True
End Sub
Sub Call_RP_BP_MOV()
    Dim strPassTry As String
    Dim lTries As Long
    Dim bSuccess As Boolean
    
    Application.ScreenUpdating = False
    For lTries = 1 To 3
        strPassTry = InputBox("Insira a Senha", "BP: Assinar & Enviar")
        If strPassTry = vbNullString Then Exit Sub
        bSuccess = strPasswordApproval = strPassTry
        If bSuccess = True Then Exit For
        MsgBox "Senha Incorreta"
    Next lTries
    If bSuccess = True Then Call BPMandaRP_MOV
    Application.ScreenUpdating = True
End Sub
Sub BPMandaRP_MOV()
    Dim Status As Boolean
    
    Status = False
    Call GetUserName_BP
    If Sheets("MOVIMENTAÇÃO").Range("W3").Value = "Sim" Then
        Status = BPMandaRP_MOV_Confidencial
    ElseIf Sheets("MOVIMENTAÇÃO").Range("W3").Value = "Não" Then
        Status = BPMandaRP_MOV_Normal
    End If
    If Status Then
        Call foo_MOV
        Call foo2_MOV
        Call MessageOK
        Call CloseCurrent
    End If
End Sub
Function BPMandaRP_MOV_Confidencial() As Boolean
    BPMandaRP_MOV_Confidencial = False
    If Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Enquadramento" Then
        BPMandaRP_MOV_Confidencial = RS_Folha_MOV_Confidencial
    ElseIf Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Movimentação Vertical" Or _
            Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Movimentação Lateral" Or _
            Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Outro" Then
        BPMandaRP_MOV_Confidencial = BP_Facilities_IT_MOV_Confidencial
        If BPMandaRP_MOV_Confidencial Then
            BPMandaRP_MOV_Confidencial = RS_Folha_MOV_Confidencial
        End If
    End If
End Function
Function BPMandaRP_MOV_Normal() As Boolean
    BPMandaRP_MOV_Normal = False
    If Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Enquadramento" Then
        BPMandaRP_MOV_Normal = RS_Folha_MOV
    ElseIf Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Movimentação Vertical" Or _
            Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Movimentação Lateral" Or _
            Sheets("MOVIMENTAÇÃO").Range("C28").Value = "Outro" Then
        BPMandaRP_MOV_Normal = BP_Facilities_IT_MOV
        BPMandaRP_MOV_Normal = RS_Folha_MOV
End If
End Function
Function Gestor_BP_MOV() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    Gestor_BP_MOV = False
    ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("MOVIMENTAÇÃO").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Solicitação de " & Range("C28") & ": " & Range("CARGOMOV")
    MailTxt = "Olá!," & vbNewLine & vbNewLine & "Segue em anexo Solicitação de " & Range("C28") & " para " & Range("CARGOMOV") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP_MOV = SendEmailMOV("MOV_GESTOR", MailTo, MailSub, MailTxt)
End Function
Function Gestor_BP_MOV_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    Gestor_BP_MOV_Confidencial = False
    ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("MOVIMENTAÇÃO").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Solicitação de " & Range("C28") & " Confidencial"
    MailTxt = "Olá!," & vbNewLine & vbNewLine & "Segue em anexo Solicitação de " & Range("C28") & " confidencial para " & Range("CARGOMOV") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP_MOV_Confidencial = SendEmailMOV("MOV_GESTOR_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT_MOV() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    BP_Facilities_IT_MOV = False
    ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailMV_BP_Facilities_IT_MOV, "MV")
    MailSub = "JML - Movimentação & Troca de Materiais/Acessos: " & Range("X11") & "_" & Range("CARGOMOV")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue abertura de chamado para Troca de Materiais & Acessos - formulário " & Range("X11") & ", referente ao cargo de " & Range("CARGOMOV") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT_MOV = SendEmailMOV("MOV_BP", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT_MOV_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    BP_Facilities_IT_MOV_Confidencial = False
    ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailMV_BP_Facilities_IT_MOV, "MV")
    MailSub = "JML - Movimentação & Troca de Materiais/Acessos Confidencial: " & Range("X11")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue abertura de chamado para Troca de Materiais & Acessos confidencial - formulário " & Range("X11") & ", referente ao funcionário " & Range("CARGOMOV") & "." & " A movimentação correspondente está sendo feita em caráter de confidencialidade e deve ser tratada com discrição." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT_MOV_Confidencial = SendEmailMOV("MOV_BP_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha_MOV() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha_MOV = False
    ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("A1").Select
    MailTo = strEmailMV_RS_Folha_MOV
    MailSub = "JML - Solicitação de " & Range("C28") & ": " & Range("X11") & "_" & Range("CARGOMOV")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue Solicitação de " & Range("C28") & " para " & Range("CARGOMOV") & "." & " O formulário " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Folha_MOV = SendEmailMOV("MOV_RS", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha_MOV_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha_MOV_Confidencial = False
    ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("A1").Select
    MailTo = strEmailMV_RS_Folha_MOV
    MailSub = "JML - Solicitação de " & Range("C28") & " Confidencial: " & Range("X11")
    MailTxt = "Olá!" & vbNewLine & vbNewLine & "Segue Solicitação de " & Range("C28") & " confidencial para " & Range("CARGOMOV") & "." & " O formulário " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Folha_MOV_Confidencial = SendEmailMOV("MOV_RS_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Sub foo_MOV()
    Dim x As Workbook
    Dim y As Workbook
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set x = ActiveWorkbook
    If strAmbiente = "producao" Then
        Set y = Workbooks.Open("\\sbra080155\rh$\BASE JML\Base Movers.xlsx")
    Else
        Set y = Workbooks.Open("\\sbra080155\public$\BI & Systems\BASE JML\Base Movers.xlsx")
    End If
    x.Sheets("MOV").Range("B4:EN4").Copy
    Range("B" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats
    Selection.PasteSpecial Paste:=xlPasteColumnWidths
    y.Save
    y.Close
End Sub
Sub foo2_MOV()
    Dim x As Workbook
    Dim y As Workbook
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set x = ActiveWorkbook
    If strAmbiente = "producao" Then
        Set y = Workbooks.Open("\\sbra080155\rh$\BASE JML - Facilities\Base Movers - Facilities.xlsx")
    Else
        Set y = Workbooks.Open("\\sbra080155\public$\BI & Systems\BASE JML - Facilities\Base Movers - Facilities.xlsx")
    End If
    x.Sheets("MOV Facilities").Range("B4:AU4").Copy
    Range("B" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats
    Selection.PasteSpecial Paste:=xlPasteColumnWidths
    y.Save
    y.Close
End Sub
Function SendEmailMOV(pType As String, pTo As String, pSubject As String, pBody As String) As Boolean
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
    'MOV 1º PASSO) MOV_GESTOR e MOV_GESTOR_CONFIDENCIAL
    'MOV 2º PASSO) MOV_BP e MOV_BP_CONFIDENCIAL
    'MOV 2º PASSO) MOV_BP_RS e MOV_BP_RS_CONFIDENCIAL
    'MOV 3º PASSO) MOV_RS e MOV_RS_CONFIDENCIAL
    'MOV 3º PASSO) MOV_RS_INFORMA e MOV_RS_INFORMA_CONFIDENCIAL
    SendEmailMOV = False
    On Error GoTo TrataErro
    With Application
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    ActiveWorkbook.Unprotect Password:=strPasswordLock
    ActiveSheet.Unprotect Password:=strPasswordLock
    TempFilePath = Environ$("temp") & "\"
    FileExtStr = "." & LCase(Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".", , 1)))
    If InStr(1, pType, "CONFIDENCIAL") > 0 Then
        TempFileNameConfidencial = " Confidencial"
    ElseIf pType = "RP_BP_CONFIDENCIAL" Then
        TempFileNameConfidencial = ""
    End If
    If pType = "MOV_GESTOR" Or pType = "MOV_GESTOR_CONFIDENCIAL" Then
        TempFileName = "JML - Solicitação de " & Range("C28") & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = ActiveWorkbook
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlRequisicaoPessoas")
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlSolicitacaoDesligamento")
        Sheets(Array("REQUISIÇÃO DE PESSOAL", "SOLICITAÇÃO DE DESLIGAMENTO", "RP", "SD", "RP Facilities", "SD Facilities")).Delete
        wb1.SaveCopyAs TempFilePath & TempFileName & FileExtStr
        On Error GoTo TrataErro
    ElseIf pType = "MOV_BP" Or pType = "MOV_BP_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Troca de Materiais e Acessos" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("MOVIMENTAÇÃO")
        With wb1
           ActiveWindow.DisplayGridlines = False
           ActiveWindow.DisplayWorkbookTabs = False
           ActiveWindow.DisplayHeadings = False
           ActiveWindow.Zoom = 90
           Set wsO = wb1.ActiveSheet
           wsI.Rows("44:118").EntireRow.Copy
           wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
           SkipBlanks:=False, Transpose:=False
           wsO.Range("A2").PasteSpecial Paste:=xlPasteColumnWidths
           wsO.Range("A2").PasteSpecial Paste:=xlPasteFormats
           ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
           ActiveSheet.EnableSelection = xlUnlockedCells
           ActiveWorkbook.Protect Structure:=True, Windows:=False
           'ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
           '    TempFileName & ".pdf", Quality:=xlQualityStandard, _
           '    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            wsO.Range("A2").Select
        End With
        wb1.SaveAs Filename:=TempFilePath & TempFileName & FileExtStr, FileFormat:=xlOpenXMLWorkbook
        wb1.Close
        On Error GoTo TrataErro
    ElseIf pType = "MOV_BP_RS" Then
        TempFileName = "JML - Solicitação de " & Range("C28") & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = ActiveWorkbook
        wb1.SaveCopyAs TempFilePath & TempFileName & FileExtStr
        On Error GoTo TrataErro
    ElseIf pType = "MOV_RS" Or pType = "MOV_RS_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Solicitação de " & Range("C28") & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("MOVIMENTAÇÃO")
        With wb1
            ActiveWindow.DisplayWorkbookTabs = False
            ActiveWindow.DisplayHeadings = False
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 90
            Set wsO = wb1.ActiveSheet
            wsI.Rows("2:43").EntireRow.Copy
            wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            wsO.Range("A2").PasteSpecial Paste:=xlPasteColumnWidths
            wsO.Range("A2").PasteSpecial Paste:=xlPasteFormats
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
            ActiveSheet.EnableSelection = xlUnlockedCells
            ActiveWorkbook.Protect Structure:=True, Windows:=False
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
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    SendEmailMOV = True
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
