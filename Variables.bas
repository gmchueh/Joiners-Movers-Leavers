Attribute VB_Name = "mdlVariables"
Option Explicit
'######################             GERAL        #######################
Public Const strPasswordApproval = "password"
Public Const strAmbiente = "producao" '"desenvolvimento"
'Email desenvolvimento
Public Const strEmailTeste = "admin1@guest.companyname.com"
'*****************************************
'           Emails de produção
'*****************************************
'######################      REQUISIÇÃO DE PESSOAS      #######################
'
Public Const strEmailRP_BP_RS = _
    "admin1@companyname.com" & ";" & _
    "admin2@guest.companyname.com"

'Public Const strEmailRP_BP_RS_Confidencial = _
'    "admin3@companyname.com" & ";" & _
'    "admin4@companyname.com"

Public Const strEmailRP_BP_Facilities_IT = _
    "admin5@companyname.com" & ";" & _
    "admin6@companyname.com" & ";" & _
    "admin7@companyname.com" & ";" & _
    "admin8@companyname.com" & ";" & _
    "admin9@companyname.com" & ";" & _
    "admin10@companyname.com"
    
'Public Const strEmailRP_BP_Facilities_IT_Confidencial = _
'    "admin11@companyname.com" & ";" & _
'    "admin12@companyname.com"

Public Const strEmailRP_RS_Folha = _
    "admin13@guest.companyname.com" & ";" & _
    "admin14@companyname.com" & ";" & _
    "admin15@companyname.com" & ";" & _
    "admin16@companyname.com"

'Public Const strEmailRP_RS_Folha_Confidencial = _
'    "admin17@guest.companyname.com"

Public Const strEmailRP_RS_Informa = _
    "admin1@companyname.com" & ";" & _
    "admin2@companyname.com" & ";" & _
    "admin3@companyname.com" & ";" & _
    "admin4@companyname.com" & ";" & _
    "admin5@companyname.com" & ";" & _
    "admin6@companyname.com" & ";" & _
    "admin7@companyname.com" & ";" & _
    "admin8@companyname.com" & ";" & _
    "admin9@companyname.com" & ";" & _
    "admin10@guest.companyname.com" & ";" & _
    "admin11@guest.companyname.com" & ";" & _
    "admin12@companyname.com" & ";" & _
    "admin13@companyname.com" & ";" & _
    "admin14@companyname.com"

'Public Const strEmailRP_RS_Informa_Confidencial = _
'    "edson.petizme@companyname.com" & ";" & _
'    "natalie.amancio@companyname.com"
'######################         MOVIMENTAÇÃO        #######################
    
Public Const strEmailMV_BP_Facilities_IT_MOV = _
    "admin1@companyname.com" & ";" & _
    "admin2@companyname.com" & ";" & _
    "admin3@companyname.com" & ";" & _
    "admin4@companyname.com" & ";" & _
    "admin5@companyname.com" & ";" & _
    "admin6@companyname.com" & ";" & _
    "admin7@companyname.com" & ";" & _
    "admin8@guest.companyname.com"

'Public Const strEmailMV_BP_Facilities_IT_MOV_Confidencial = _
'    "admin1@companyname.com" & ";" & _
'    "admin2@companyname.com" & ";" & _
'    "admin3@companyname.com"
    
Public Const strEmailMV_RS_Folha_MOV = _
    "admin@guest.companyname.com" & ";" & _
    "admin@companyname.com" & ";" & _
    "admin@companyname.com" & ";" & _
    "admin@companyname.com" & ";"

'Public Const strEmailMV_RS_Folha_MOV_Confidencial = _
'    "admin@guest.companyname.com"

'######################  SOLICITAÇÃO DE DESLIGAMENTO #######################
Public Const strEmailSD_RS_Folha_SD = _
    "admin@guest.companyname.com" & ";" & _
    "admin@companyname.com" & ";" & _
    "admin@companyname.com" & ";" & _
    "admin@companyname.com" & ";"
    
'Public Const strEmailSD_RS_Folha_SD_Confidencial = _
'    "admin1@guest.companyname.com" & ";" & _
'    "admin2@companyname.com"
    
Public Const strEmailSD_BP_Facilities_IT_SD = _
    "admin1@companyname.com" & ";" & _
    "admin2@companyname.com" & ";" & _
    "admin3@companyname.com" & ";" & _
    "admin4@companyname.com" & ";" & _
    "admin5@companyname.com" & ";" & _
    "admin6@companyname.com" & ";" & _
    "admin7@companyname.com" & ";" & _
    "admin8@companyname.com" & ";" & _
    "admin9@guest.companyname.com"

'Public Const strEmailSD_BP_Facilities_IT_SD_Confidencial = _
'    "admin1@companyname.com" & ";" & _
'    "admin1@companyname.com"
    
'############################    TESTE          #############################
'*****************************************
'           Emails de Testes
'*****************************************
'RP
'Public Const strEmailRP_BP_Facilities_IT = strEmailTeste
'Public Const strEmailRP_BP_Facilities_IT_Confidencial = strEmailTeste
'Public Const strEmailRP_BP_RS = strEmailTeste
'Public Const strEmailRP_BP_RS_Confidencial = strEmailTeste
'Public Const strEmailRP_RS_Folha = strEmailTeste
'Public Const strEmailRP_RS_Folha_Confidencial = strEmailTeste
'Public Const strEmailRP_RS_Informa = strEmailTeste
'Public Const strEmailRP_RS_Informa_Confidencial = strEmailTeste
'MOV
'Public Const strEmailMV_BP_Facilities_IT_MOV = strEmailTeste
'Public Const strEmailMV_BP_Facilities_IT_MOV_Confidencial = strEmailTeste
'Public Const strEmailMV_BP_RS_MOV = strEmailTeste
'Public Const strEmailMV_RS_Folha_MOV = strEmailTeste
'Public Const strEmailMV_RS_Folha_MOV_Confidencial = strEmailTeste
'SDE
'Public Const strEmailSD_BP_Facilities_IT_SD = strEmailTeste
'Public Const strEmailSD_BP_Facilities_IT_SD_Confidencial = strEmailTeste
'Public Const strEmailSD_RS_Folha_SD = strEmailTeste
'Public Const strEmailSD_RS_Folha_SD_Confidencial = strEmailTeste

Public Const strPasswordLock = "companynameTI"

Sub GetUserName_Gestor()
    Range("S9").Value = Environ("Username")
End Sub
Sub GetUserName_BP()
    Range("S11").Value = Environ("Username")
End Sub
Sub CloseCurrent()
    ActiveWorkbook.Close False
End Sub
Sub MessageOK()
    MsgBox "Enviado com sucesso!"
End Sub
Sub ClearUnlockedCells()
    Application.FindFormat.Clear
    Application.FindFormat.Locked = False
    Cells.Replace "*", "", SearchFormat:=True
    Application.FindFormat.Clear
End Sub
Function VerificaEnvioIT(listaEmail As String, tipoEnvio) As String
    VerificaEnvioIT = listaEmail
    
    If (tipoEnvio = "RP") Then
        If ((UCase(ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("CARGORP")) = "PROMOTOR" Or UCase(ThisWorkbook.Sheets("REQUISIÇÃO DE PESSOAL").Range("CARGORP")) = "PROMOTOR PANELISTA")) Then
            VerificaEnvioIT = Replace(VerificaEnvioIT, "tkt.brazil.support@companyname.com;", "")
            VerificaEnvioIT = Replace(VerificaEnvioIT, "edson.petizme@companyname.com;", "")
            VerificaEnvioIT = Replace(VerificaEnvioIT, "DLL-BRA08-ITS-Team@companyname.com;", "")
        End If
    ElseIf (tipoEnvio = "MV") Then
        If ((UCase(ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("C19")) = "PROMOTOR" Or UCase(ThisWorkbook.Sheets("MOVIMENTAÇÃO").Range("C19")) = "PROMOTOR PANELISTA")) Then
            VerificaEnvioIT = Replace(VerificaEnvioIT, "tkt.brazil.support@companyname.com;", "")
            VerificaEnvioIT = Replace(VerificaEnvioIT, "edson.petizme@companyname.com;", "")
            VerificaEnvioIT = Replace(VerificaEnvioIT, "DLL-BRA08-ITS-Team@companyname.com;", "")
        End If
    ElseIf (tipoEnvio = "SD") Then
        If ((UCase(ThisWorkbook.Sheets("SOLICITAÇÃO DE DESLIGAMENTO").Range("CARGORD")) = "PROMOTOR" Or UCase(ThisWorkbook.Sheets("SOLICITAÇÃO DE DESLIGAMENTO").Range("CARGORD")) = "PROMOTOR PANELISTA")) Then
            VerificaEnvioIT = Replace(VerificaEnvioIT, "tkt.brazil.support@companyname.com;", "")
            VerificaEnvioIT = Replace(VerificaEnvioIT, "edson.petizme@companyname.com;", "")
            VerificaEnvioIT = Replace(VerificaEnvioIT, "DLL-BRA08-ITS-Team@companyname.com;", "")
        End If
    End If
    
End Function
'Remove acentos e caracteres especiais e deixa o texto em caixa alta
Public Function RetirarCaracteres(ByVal Caract As Variant) As Variant
'Declaracao de Variaveis
Dim i       As Long
Dim p       As Variant
Dim codiA   As String
Dim codiB   As String

'Caracteres impeditivos
codiA = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ-'´)([]/\*-+.,!@#$%¨&§¹²³£¢¬"
'Caracteres substitutivos
codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN____________________________"
    
    'Inicia o loop em busca dos caracteres impeditivos
    For i = 1 To Len(Caract)
        p = InStr(codiA, Mid(Caract, i, 1))
        'Verifica a existencia dos caracteres no texto
        If p > 0 Then
            'Realiza a substituicao
            Mid(Caract, i, 1) = Mid(codiB, p, 1)
        End If
    Next

'Retorno do texto
RetirarCaracteres = Application.WorksheetFunction.Trim(Caract)
     
End Function
