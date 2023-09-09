Attribute VB_Name = "mdlMovimentacao"
Option Explicit
Sub Obrigatorio_MOV()
    If _
        Sheets("MOVIMENTA��O").Range("CARGOMOV").Value = "" Or Sheets("MOVIMENTA��O").Range("CCATUAL").Value = "" Or Sheets("MOVIMENTA��O").Range("W3").Value = "" Or Sheets("MOVIMENTA��O").Range("J7").Value = "" Or Sheets("MOVIMENTA��O").Range("N7").Value = "" Or Sheets("MOVIMENTA��O").Range("Q7").Value = "" Or Sheets("MOVIMENTA��O").Range("J11").Value = "" Or Sheets("MOVIMENTA��O").Range("P11").Value = "" Or Sheets("MOVIMENTA��O").Range("C16").Value = "" Or Sheets("MOVIMENTA��O").Range("J16").Value = "" Or Sheets("MOVIMENTA��O").Range("Q16").Value = "" Or Sheets("MOVIMENTA��O").Range("C19").Value = "" Or _
        Sheets("MOVIMENTA��O").Range("J19").Value = "" Or Sheets("MOVIMENTA��O").Range("Q19").Value = "" Or Sheets("MOVIMENTA��O").Range("C22").Value = "" Or Sheets("MOVIMENTA��O").Range("J22").Value = "" Or Sheets("MOVIMENTA��O").Range("O22").Value = "" Or Sheets("MOVIMENTA��O").Range("Q22").Value = "" Or Sheets("MOVIMENTA��O").Range("C25").Value = "" Or Sheets("MOVIMENTA��O").Range("J25").Value = "" Or Sheets("MOVIMENTA��O").Range("O25").Value = "" Or Sheets("MOVIMENTA��O").Range("Q25").Value = "" Or Sheets("MOVIMENTA��O").Range("C28").Value = "" Or Sheets("MOVIMENTA��O").Range("J28").Value = "" Or _
        Sheets("MOVIMENTA��O").Range("Q28").Value = "" Or Sheets("MOVIMENTA��O").Range("V28").Value = "" Or Sheets("MOVIMENTA��O").Range("V57").Value = "" Or Sheets("MOVIMENTA��O").Range("C66").Value = "" Or Sheets("MOVIMENTA��O").Range("E66").Value = "" Or Sheets("MOVIMENTA��O").Range("G66").Value = "" Or Sheets("MOVIMENTA��O").Range("K66").Value = "" Or Sheets("MOVIMENTA��O").Range("U66").Value = "" Or Sheets("MOVIMENTA��O").Range("C69").Value = "" Or Sheets("MOVIMENTA��O").Range("E69").Value = "" Or Sheets("MOVIMENTA��O").Range("G69").Value = "" Or Sheets("MOVIMENTA��O").Range("K69").Value = "" Or _
        Sheets("MOVIMENTA��O").Range("U69").Value = "" Or Sheets("MOVIMENTA��O").Range("T96").Value = "" Or Sheets("MOVIMENTA��O").Range("V96").Value = "" Or Sheets("MOVIMENTA��O").Range("T97").Value = "" Or Sheets("MOVIMENTA��O").Range("V97").Value = "" Or Sheets("MOVIMENTA��O").Range("T98").Value = "" Or Sheets("MOVIMENTA��O").Range("V98").Value = "" Or Sheets("MOVIMENTA��O").Range("T99").Value = "" Or Sheets("MOVIMENTA��O").Range("V99").Value = "" Or Sheets("MOVIMENTA��O").Range("T100").Value = "" Or Sheets("MOVIMENTA��O").Range("V100").Value = "" Or Sheets("MOVIMENTA��O").Range("T101").Value = "" Or _
        Sheets("MOVIMENTA��O").Range("V101").Value = "" Or Sheets("MOVIMENTA��O").Range("T102").Value = "" Or Sheets("MOVIMENTA��O").Range("V102").Value = "" Or Sheets("MOVIMENTA��O").Range("T103").Value = "" Or Sheets("MOVIMENTA��O").Range("V103").Value = "" Or Sheets("MOVIMENTA��O").Range("T104").Value = "" Or Sheets("MOVIMENTA��O").Range("V104").Value = "" Or Sheets("MOVIMENTA��O").Range("T105").Value = "" Or Sheets("MOVIMENTA��O").Range("V105").Value = "" Or Sheets("MOVIMENTA��O").Range("T106").Value = "" Or Sheets("MOVIMENTA��O").Range("V106").Value = "" Or Sheets("MOVIMENTA��O").Range("T107").Value = "" Or _
        Sheets("MOVIMENTA��O").Range("V107").Value = "" Or Sheets("MOVIMENTA��O").Range("T108").Value = "" Or Sheets("MOVIMENTA��O").Range("V108").Value = "" Or Sheets("MOVIMENTA��O").Range("T109").Value = "" Or Sheets("MOVIMENTA��O").Range("V109").Value = "" Or Sheets("MOVIMENTA��O").Range("T110").Value = "" Or Sheets("MOVIMENTA��O").Range("V110").Value = "" Or Sheets("MOVIMENTA��O").Range("T111").Value = "" Or Sheets("MOVIMENTA��O").Range("V111").Value = "" Or Sheets("MOVIMENTA��O").Range("T112").Value = "" Or Sheets("MOVIMENTA��O").Range("V112").Value = "" Or Sheets("MOVIMENTA��O").Range("T113").Value = "" Or Sheets("MOVIMENTA��O").Range("V113").Value = "" _
    Then
        MsgBox "Obrigat�rio o preenchimento de todos os campos em Vermelho"
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
    If Sheets("MOVIMENTA��O").Range("W3").Value = "N�o" Then
        Status = Gestor_BP_MOV
    ElseIf Sheets("MOVIMENTA��O").Range("W3").Value = "Sim" Then
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
    If Sheets("MOVIMENTA��O").Range("W3").Value = "Sim" Then
        Status = BPMandaRP_MOV_Confidencial
    ElseIf Sheets("MOVIMENTA��O").Range("W3").Value = "N�o" Then
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
    If Sheets("MOVIMENTA��O").Range("C28").Value = "Enquadramento" Then
        BPMandaRP_MOV_Confidencial = RS_Folha_MOV_Confidencial
    ElseIf Sheets("MOVIMENTA��O").Range("C28").Value = "Movimenta��o Vertical" Or _
            Sheets("MOVIMENTA��O").Range("C28").Value = "Movimenta��o Lateral" Or _
            Sheets("MOVIMENTA��O").Range("C28").Value = "Outro" Then
        BPMandaRP_MOV_Confidencial = BP_Facilities_IT_MOV_Confidencial
        If BPMandaRP_MOV_Confidencial Then
            BPMandaRP_MOV_Confidencial = RS_Folha_MOV_Confidencial
        End If
    End If
End Function
Function BPMandaRP_MOV_Normal() As Boolean
    BPMandaRP_MOV_Normal = False
    If Sheets("MOVIMENTA��O").Range("C28").Value = "Enquadramento" Then
        BPMandaRP_MOV_Normal = RS_Folha_MOV
    ElseIf Sheets("MOVIMENTA��O").Range("C28").Value = "Movimenta��o Vertical" Or _
            Sheets("MOVIMENTA��O").Range("C28").Value = "Movimenta��o Lateral" Or _
            Sheets("MOVIMENTA��O").Range("C28").Value = "Outro" Then
        BPMandaRP_MOV_Normal = BP_Facilities_IT_MOV
        BPMandaRP_MOV_Normal = RS_Folha_MOV
End If
End Function
Function Gestor_BP_MOV() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    Gestor_BP_MOV = False
    ThisWorkbook.Sheets("MOVIMENTA��O").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("MOVIMENTA��O").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Solicita��o de " & Range("C28") & ": " & Range("CARGOMOV")
    MailTxt = "Ol�!," & vbNewLine & vbNewLine & "Segue em anexo Solicita��o de " & Range("C28") & " para " & Range("CARGOMOV") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP_MOV = SendEmailMOV("MOV_GESTOR", MailTo, MailSub, MailTxt)
End Function
Function Gestor_BP_MOV_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    Gestor_BP_MOV_Confidencial = False
    ThisWorkbook.Sheets("MOVIMENTA��O").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("MOVIMENTA��O").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Solicita��o de " & Range("C28") & " Confidencial"
    MailTxt = "Ol�!," & vbNewLine & vbNewLine & "Segue em anexo Solicita��o de " & Range("C28") & " confidencial para " & Range("CARGOMOV") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP_MOV_Confidencial = SendEmailMOV("MOV_GESTOR_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT_MOV() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    BP_Facilities_IT_MOV = False
    ThisWorkbook.Sheets("MOVIMENTA��O").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailMV_BP_Facilities_IT_MOV, "MV")
    MailSub = "JML - Movimenta��o & Troca de Materiais/Acessos: " & Range("X11") & "_" & Range("CARGOMOV")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue abertura de chamado para Troca de Materiais & Acessos - formul�rio " & Range("X11") & ", referente ao cargo de " & Range("CARGOMOV") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT_MOV = SendEmailMOV("MOV_BP", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT_MOV_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String

    BP_Facilities_IT_MOV_Confidencial = False
    ThisWorkbook.Sheets("MOVIMENTA��O").Range("A1").Select
    MailTo = VerificaEnvioIT(strEmailMV_BP_Facilities_IT_MOV, "MV")
    MailSub = "JML - Movimenta��o & Troca de Materiais/Acessos Confidencial: " & Range("X11")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue abertura de chamado para Troca de Materiais & Acessos confidencial - formul�rio " & Range("X11") & ", referente ao funcion�rio " & Range("CARGOMOV") & "." & " A movimenta��o correspondente est� sendo feita em car�ter de confidencialidade e deve ser tratada com discri��o." & vbNewLine & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT_MOV_Confidencial = SendEmailMOV("MOV_BP_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha_MOV() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha_MOV = False
    ThisWorkbook.Sheets("MOVIMENTA��O").Range("A1").Select
    MailTo = strEmailMV_RS_Folha_MOV
    MailSub = "JML - Solicita��o de " & Range("C28") & ": " & Range("X11") & "_" & Range("CARGOMOV")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue Solicita��o de " & Range("C28") & " para " & Range("CARGOMOV") & "." & " O formul�rio " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Folha_MOV = SendEmailMOV("MOV_RS", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha_MOV_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha_MOV_Confidencial = False
    ThisWorkbook.Sheets("MOVIMENTA��O").Range("A1").Select
    MailTo = strEmailMV_RS_Folha_MOV
    MailSub = "JML - Solicita��o de " & Range("C28") & " Confidencial: " & Range("X11")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue Solicita��o de " & Range("C28") & " confidencial para " & Range("CARGOMOV") & "." & " O formul�rio " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
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
    
    'Os tipos de forma de envio de email (pType) s�o:
    'MOV 1� PASSO) MOV_GESTOR e MOV_GESTOR_CONFIDENCIAL
    'MOV 2� PASSO) MOV_BP e MOV_BP_CONFIDENCIAL
    'MOV 2� PASSO) MOV_BP_RS e MOV_BP_RS_CONFIDENCIAL
    'MOV 3� PASSO) MOV_RS e MOV_RS_CONFIDENCIAL
    'MOV 3� PASSO) MOV_RS_INFORMA e MOV_RS_INFORMA_CONFIDENCIAL
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
        TempFileName = "JML - Solicita��o de " & Range("C28") & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = ActiveWorkbook
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlRequisicaoPessoas")
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlSolicitacaoDesligamento")
        Sheets(Array("REQUISI��O DE PESSOAL", "SOLICITA��O DE DESLIGAMENTO", "RP", "SD", "RP Facilities", "SD Facilities")).Delete
        wb1.SaveCopyAs TempFilePath & TempFileName & FileExtStr
        On Error GoTo TrataErro
    ElseIf pType = "MOV_BP" Or pType = "MOV_BP_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Troca de Materiais e Acessos" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("MOVIMENTA��O")
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
        TempFileName = "JML - Solicita��o de " & Range("C28") & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = ActiveWorkbook
        wb1.SaveCopyAs TempFilePath & TempFileName & FileExtStr
        On Error GoTo TrataErro
    ElseIf pType = "MOV_RS" Or pType = "MOV_RS_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Solicita��o de " & Range("C28") & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGOMOV") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("MOVIMENTA��O")
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
    MsgBox "Ocorreu um erro ao enviar o email. Por favor, contate o administrador. Descri��o : " & Err.Description
End Function
