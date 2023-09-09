Attribute VB_Name = "mdlSolicitacaoDesligamento"
Option Explicit
Sub Obrigatorio_SD()
    If ((InStr(1, Sheets("SOLICITA��O DE DESLIGAMENTO").Range("CARGORD").Value, "DIRETOR") > 0 Or _
         InStr(1, Sheets("SOLICITA��O DE DESLIGAMENTO").Range("CARGORD").Value, "GERENTE") > 0 Or _
         InStr(1, Sheets("SOLICITA��O DE DESLIGAMENTO").Range("CARGORD").Value, "GER REG") > 0) And _
         ((Sheets("SOLICITA��O DE DESLIGAMENTO").Range("C50").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("C50").Value = "Sim")) And _
          Sheets("SOLICITA��O DE DESLIGAMENTO").Range("G50").Value = "") _
    Then
        MsgBox "Para o cargo de " & Sheets("SOLICITA��O DE DESLIGAMENTO").Range("CARGORD").Value & " � obrigat�rio informar a necessidade de realizar backup dos arquivos e o funcion�rio que ficar� com o backup. Estes campos est�o no final da planilha."
        Exit Sub
    ElseIf (Sheets("SOLICITA��O DE DESLIGAMENTO").Range("C50").Value = "Sim" And _
             Sheets("SOLICITA��O DE DESLIGAMENTO").Range("G50").Value = "") _
    Then
        MsgBox "Informe o funcion�rio que ficar� com o backup. Este campo est� no final da planilha."
        Exit Sub
    ElseIf _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("CARGORD").Value = "" And Sheets("SOLICITA��O DE DESLIGAMENTO").Range("CCATUA").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("K7").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("N7").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("Q7").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("C16").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("W3").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("k11").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("j16").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("m16").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("r16").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("j19").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("n19").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("c22").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("k47").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("j22").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("n47").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("k44").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("n44").Value = "" Or _
        Sheets("SOLICITA��O DE DESLIGAMENTO").Range("c47").Value = "" Or Sheets("SOLICITA��O DE DESLIGAMENTO").Range("g47").Value = "" _
    Then
        MsgBox "Obrigat�rio o preenchimento de todos os campos em Vermelho"
        Exit Sub
    Else
        Call GestorMandaRP_SD
    End If
End Sub
Sub GestorMandaRP_SD()
    Dim Status As Boolean
    
    Status = False
    Application.ScreenUpdating = False
    Call GetUserName_Gestor
    If Sheets("SOLICITA��O DE DESLIGAMENTO").Range("W3").Value = "N�o" Then
        Status = Gestor_BP_SD
    ElseIf Sheets("SOLICITA��O DE DESLIGAMENTO").Range("W3").Value = "Sim" Then
        Status = Gestor_BP_SD_Confidencial
    End If
    If Status Then
        'Call ClearUnlockedCells
        Call MessageOK
        Call CloseCurrent
    End If
    Application.ScreenUpdating = True
End Sub
Sub Call_RP_BP_SD()
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
    If bSuccess = True Then Call BPMandaRP_SD
    Application.ScreenUpdating = True
End Sub
Sub BPMandaRP_SD()
    Dim Status As Boolean
    
    Status = False
    Call GetUserName_BP
    If Sheets("SOLICITA��O DE DESLIGAMENTO").Range("W3").Value = "N�o" Then
        Status = BP_Facilities_IT_SD
        If Status Then
            Status = RS_Folha_SD
        End If
    ElseIf Sheets("SOLICITA��O DE DESLIGAMENTO").Range("W3").Value = "Sim" Then
        Status = BP_Facilities_IT_SD_Confidencial
        If Status Then
            Status = RS_Folha_SD_Confidencial
        End If
    End If
    If Status Then
        Call foo_SD
        Call foo2_SD
        Call MessageOK
        Call CloseCurrent
    End If
End Sub
Function Gestor_BP_SD() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    Gestor_BP_SD = False
    ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("SOLICITA��O DE DESLIGAMENTO").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Solicita��o de Desligamento: " & Range("CARGORD") & "_" & Range("N7") & "_" & Range("Q7")
    MailTxt = "Ol�!," & vbNewLine & vbNewLine & "Segue em anexo Solicita��o de Desligamento para " & Range("CARGORP") & " " & "em" & " " & Range("N7") & "/" & Range("Q7") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP_SD = SendEmailSD("SD_GESTOR", MailTo, MailSub, MailTxt)
End Function
Function Gestor_BP_SD_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    Gestor_BP_SD_Confidencial = False
    ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = Sheets("SOLICITA��O DE DESLIGAMENTO").Range("email")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Solicita��o de Desligamento confidencial"
    MailTxt = "Ol�!," & vbNewLine & vbNewLine & "Segue em anexo Solicita��o de Desligamento confidencial para " & Range("CARGORP") & " " & "em" & " " & Range("N7") & "/" & Range("Q7") & "." & vbNewLine & vbNewLine & "Atenciosamente,"
    Gestor_BP_SD_Confidencial = SendEmailSD("SD_GESTOR_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT_SD() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    BP_Facilities_IT_SD = False
    ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = VerificaEnvioIT(strEmailSD_BP_Facilities_IT_SD, "SD")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Desligamento: " & Range("X11") & "_" & Range("CARGORD") & "_" & Range("N7") & "_" & Range("Q7")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue abertura de chamado para Devolu��o de Materiais & Acessos, formul�rio " & Range("X11") & " referente ao desligamento do(a) " & Range("CARGORD") & " " & "de" & " " & Range("N7") & " " & Range("Q7") & "." & vbNewLine
    If Sheets("SOLICITA��O DE DESLIGAMENTO").Range("C50").Value = "Sim" Then
        MailTxt = MailTxt & vbNewLine & "ATEN��O: FOI SOLICITADA A GERA��O DO BACKUP DOS ARQUIVOS DO COLABORADOR. MAIORES DETALHES NA JML EM ANEXO" & vbNewLine
    End If
    MailTxt = MailTxt & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT_SD = SendEmailSD("SD_BP", MailTo, MailSub, MailTxt)
End Function
Function BP_Facilities_IT_SD_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    BP_Facilities_IT_SD_Confidencial = False
    ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO").Range("A1").Select
    If strAmbiente = "producao" Then
        MailTo = VerificaEnvioIT(strEmailSD_BP_Facilities_IT_SD, "SD")
    Else
        MailTo = strEmailTeste
    End If
    MailSub = "JML - Desligamento Confidencial: " & Range("X11")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue abertura de chamado para Devolu��o de Materiais & Acessos confidencial, formul�rio " & Range("X11") & ", referente ao desligamento do(a) " & Range("CARGORD") & " " & "de" & " " & Range("N7") & " " & Range("Q7") & "." & " O desligamento correspondente est� sendo feito em car�ter de confidencialidade e deve ser tratado com discri��o." & vbNewLine
    If Sheets("SOLICITA��O DE DESLIGAMENTO").Range("C50").Value = "Sim" Then
        MailTxt = MailTxt & vbNewLine & "ATEN��O: FOI SOLICITADA A GERA��O DO BACKUP DOS ARQUIVOS DO COLABORADOR. MAIORES DETALHES NA JML EM ANEXO" & vbNewLine
    End If
    MailTxt = MailTxt & vbNewLine & "Atenciosamente,"
    BP_Facilities_IT_SD_Confidencial = SendEmailSD("SD_BP_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha_SD() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha_SD = False
    ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO").Range("A1").Select
    MailTo = strEmailSD_RS_Folha_SD
    MailSub = "JML - Solicita��o de Desligamento: " & Range("X11") & "_" & Range("CARGORD") & "_" & Range("N7") & "_" & Range("Q7")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue Solicita��o de Desligamento para " & Range("CARGORD") & " " & "de" & " " & Range("N7") & "-" & Range("Q7") & "." & " O formul�rio " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Folha_SD = SendEmailSD("SD_RS", MailTo, MailSub, MailTxt)
End Function
Function RS_Folha_SD_Confidencial() As Boolean
    Dim MailTo As String, MailSub As String, MailTxt As String
    
    RS_Folha_SD_Confidencial = False
    ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO").Range("A1").Select
    MailTo = strEmailSD_RS_Folha_SD
    MailSub = "JML - Solicita��o de Desligamento Confidencial: " & Range("X11")
    MailTxt = "Ol�!" & vbNewLine & vbNewLine & "Segue Solicita��o de Desligamento confidencial para " & Range("CARGORD") & " " & "de" & " " & Range("N7") & "-" & Range("Q7") & "." & " O formul�rio " & Range("X11") & " em anexo contempla todos os detalhes." & vbNewLine & vbNewLine & "Atenciosamente,"
    RS_Folha_SD_Confidencial = SendEmailSD("SD_RS_CONFIDENCIAL", MailTo, MailSub, MailTxt)
End Function
Sub foo_SD()
    Dim x As Workbook
    Dim y As Workbook
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set x = ActiveWorkbook
    If strAmbiente = "producao" Then
        Set y = Workbooks.Open("\\sbra080155\rh$\BASE JML\Base Leavers.xlsx")
    Else
        Set y = Workbooks.Open("\\sbra080155\public$\BI & Systems\BASE JML\Base Leavers.xlsx")
    End If
    x.Sheets("SD").Range("B4:Z4").Copy
    Range("B" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats
    Selection.PasteSpecial Paste:=xlPasteColumnWidths
    y.Save
    y.Close
End Sub
Sub foo2_SD()
    Dim x As Workbook
    Dim y As Workbook
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set x = ActiveWorkbook
    If strAmbiente = "producao" Then
        Set y = Workbooks.Open("\\sbra080155\rh$\BASE JML - Facilities\Base Leavers - Facilities.xlsx")
    Else
        Set y = Workbooks.Open("\\sbra080155\public$\BI & Systems\BASE JML - Facilities\Base Leavers - Facilities.xlsx")
    End If
    x.Sheets("SD Facilities").Range("B4:AH4").Copy
    Range("B" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats
    Selection.PasteSpecial Paste:=xlPasteColumnWidths
    y.Save
    y.Close
End Sub
Function SendEmailSD(pType As String, pTo As String, pSubject As String, pBody As String) As Boolean
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
    'SD 1� PASSO) SD_GESTOR e SD_GESTOR_CONFIDENCIAL
    'SD 2� PASSO) SD_BP e SD_BP_CONFIDENCIAL
    'SD 2� PASSO) SD_BP_RS e SD_BP_RS_CONFIDENCIAL
    SendEmailSD = False
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
    If pType = "SD_GESTOR" Or pType = "SD_GESTOR_CONFIDENCIAL" Then
        TempFileName = "JML - Solicita��o de Desligamento" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("CARGORD") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = ActiveWorkbook
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlRequisicaoPessoas")
        wb1.VBProject.VBComponents.Remove wb1.VBProject.VBComponents("mdlMovimentacao")
        Sheets(Array("REQUISI��O DE PESSOAL", "MOVIMENTA��O", "RP", "MOV", "RP Facilities", "Mov Facilities")).Delete
        wb1.SaveCopyAs TempFilePath & TempFileName & FileExtStr
        On Error GoTo TrataErro
    ElseIf pType = "SD_BP" Or pType = "SD_BP_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Devolu��o de Materiais & Acessos" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGORD") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO")
        With wb1
           ActiveWindow.DisplayGridlines = False
           ActiveWindow.DisplayWorkbookTabs = False
           ActiveWindow.DisplayHeadings = False
           ActiveWindow.Zoom = 90
           Set wsO = wb1.ActiveSheet
           wsI.Rows("34:59").EntireRow.Copy
           wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
           SkipBlanks:=False, Transpose:=False
           wsO.Range("A2").PasteSpecial Paste:=xlPasteColumnWidths
           wsO.Range("A2").PasteSpecial Paste:=xlPasteFormats
           ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
           ActiveSheet.EnableSelection = xlUnlockedCells
           ActiveWorkbook.Protect Structure:=True, Windows:=False
           wsO.Range("A2").Select
        End With
        wb1.SaveAs Filename:=TempFilePath & TempFileName & FileExtStr, FileFormat:=xlOpenXMLWorkbook
        wb1.Close
        On Error GoTo TrataErro
    ElseIf pType = "SD_RS" Or pType = "SD_RS_CONFIDENCIAL" Then
        FileExtStr = ".xlsx"
        TempFileName = "JML - Solicita��o de Desligamento" & TempFileNameConfidencial & "_" & RetirarCaracteres(Range("X11") & "_" & Range("CARGORD") & "_" & Range("N7") & "_" & Range("Q7"))
        On Error Resume Next
        Set wb1 = Workbooks.Add
        Set wsI = ThisWorkbook.Sheets("SOLICITA��O DE DESLIGAMENTO")
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
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
            ActiveSheet.EnableSelection = xlUnlockedCells
            ActiveWorkbook.Protect Structure:=True, Windows:=False
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
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    SendEmailSD = True
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


