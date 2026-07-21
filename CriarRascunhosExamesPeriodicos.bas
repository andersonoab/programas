Attribute VB_Name = "ModExamesPeriodicos"
Option Explicit

'==============================================================================
' MACRO: CriarRascunhosExamesPeriodicos
'------------------------------------------------------------------------------
' OBJETIVO
'   Ler a planilha de pendencias de exames ocupacionais (1 linha por
'   colaborador/exame), agrupar por GESTOR e criar, no Outlook, UM RASCUNHO
'   por gestor com a lista da equipe, orientando a organizacao da agenda:
'   exames marcados na clinica mais proxima do trabalho, sempre pela manha,
'   com atendimento ate as 11h. Contexto: inspecao ANVISA / ponto de auditoria.
'
' SAIDA
'   Rascunhos salvos na pasta "Rascunhos" do Outlook. NADA E ENVIADO.
'   Linhas com gestor vazio ou em erro (#N/D, #NOME? etc.) geram um rascunho
'   separado de pendencias internas, sem destinatario.
'
' COMO EXECUTAR
'   1) Abra a planilha (aba com os dados ativa) e o Outlook.
'   2) ALT+F11 > Inserir > Modulo > cole este codigo.
'   3) F5 (ou ALT+F8) em CriarRascunhosExamesPeriodicos.
'
' REQUISITOS
'   - Outlook aberto com a conta corporativa.
'   - Coluna "Gestor" com valores calculados (se vier de PROCX com vinculo
'     externo, abrir com vinculos atualizados; linhas em erro viram pendencia).
'   - Coluna OPCIONAL "E-mail Gestor": se existir, e usada como destinatario.
'     Se nao existir, a macro tenta resolver o NOME do gestor no catalogo de
'     enderecos (GAL). O que nao resolver fica como texto no campo Para, para
'     ajuste manual antes do envio (sao rascunhos).
'
' RH Sonova - Anderson Souza
'==============================================================================

'----------------------------- CONFIGURACOES ----------------------------------
Private Const NOME_ABA As String = ""       ' "" = usa a aba ativa
Private Const LINHA_CABECALHO As Long = 1   ' linha dos titulos das colunas
Private Const CC_FIXO As String = ""        ' opcional, ex.: "sst@empresa.com"
Private Const CRIAR_RASCUNHO_SEM_GESTOR As Boolean = True

' Cabecalhos esperados (comparacao sem acentos e sem maiusculas/minusculas)
Private Const H_FUNC As String = "funcionario"
Private Const H_MAT As String = "matricula"
Private Const H_GESTOR As String = "gestor"
Private Const H_LOCAL As String = "local de trabalho"
Private Const H_UF As String = "local de trabalho (estado)"
Private Const H_CIDADE As String = "local de trabalho (cidade)"
Private Const H_EXAME As String = "exame"
Private Const H_SIT As String = "situacao exame"
Private Const H_PROX As String = "data proximo"
Private Const H_EMAIL As String = "e-mail gestor"   ' coluna opcional

'==============================================================================
Public Sub CriarRascunhosExamesPeriodicos()

    Dim ws As Worksheet
    Dim dictCols As Object, dictGrupos As Object
    Dim dictNomes As Object, dictEmails As Object
    Dim olApp As Object, olMail As Object, olInsp As Object, olRcp As Object
    Dim r As Long, ultimaLinha As Long
    Dim chave As String, gestorBruto As String, nomeExib As String
    Dim itens As Collection
    Dim qtdRascunhos As Long, qtdColabs As Long, qtdSemGestor As Long
    Dim assinatura As String, emailDest As String
    Dim vGestor As Variant

    '--- Planilha de origem -----------------------------------------------
    If NOME_ABA = "" Then
        Set ws = ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    End If

    '--- Mapear colunas pelo cabecalho ------------------------------------
    Set dictCols = MapearColunas(ws)
    If Not ValidarColunas(dictCols) Then Exit Sub

    ultimaLinha = ws.Cells(ws.Rows.Count, dictCols(H_FUNC)).End(xlUp).Row
    If ultimaLinha <= LINHA_CABECALHO Then
        MsgBox "Nenhuma linha de dados encontrada abaixo do cabecalho.", vbExclamation
        Exit Sub
    End If

    '--- Agrupar por gestor ------------------------------------------------
    Set dictGrupos = CreateObject("Scripting.Dictionary")
    Set dictNomes = CreateObject("Scripting.Dictionary")
    Set dictEmails = CreateObject("Scripting.Dictionary")

    For r = LINHA_CABECALHO + 1 To ultimaLinha
        If TextoCel(ws.Cells(r, dictCols(H_FUNC)).Value) <> "" Then

            vGestor = ws.Cells(r, dictCols(H_GESTOR)).Value
            If IsError(vGestor) Then
                gestorBruto = ""
            Else
                gestorBruto = Trim$(CStr(vGestor))
            End If

            If gestorBruto = "" Then
                chave = "(SEM GESTOR IDENTIFICADO)"
                nomeExib = chave
                qtdSemGestor = qtdSemGestor + 1
            Else
                chave = UCase$(gestorBruto)
                nomeExib = StrConv(gestorBruto, vbProperCase)
            End If

            If Not dictGrupos.Exists(chave) Then
                dictGrupos.Add chave, New Collection
                dictNomes.Add chave, nomeExib
                ' E-mail do gestor (coluna opcional)
                If dictCols.Exists(H_EMAIL) Then
                    dictEmails.Add chave, TextoCel(ws.Cells(r, dictCols(H_EMAIL)).Value)
                Else
                    dictEmails.Add chave, ""
                End If
            End If

            dictGrupos(chave).Add Array( _
                TextoCel(ws.Cells(r, dictCols(H_FUNC)).Value), _
                TextoCel(ws.Cells(r, dictCols(H_MAT)).Value), _
                TextoCel(ws.Cells(r, dictCols(H_LOCAL)).Value), _
                MontarCidadeUF(TextoCel(ws.Cells(r, dictCols(H_CIDADE)).Value), _
                               TextoCel(ws.Cells(r, dictCols(H_UF)).Value)), _
                TextoCel(ws.Cells(r, dictCols(H_EXAME)).Value), _
                TextoCel(ws.Cells(r, dictCols(H_SIT)).Value), _
                FormatarDataBR(ws.Cells(r, dictCols(H_PROX)).Value))

            qtdColabs = qtdColabs + 1
        End If
    Next r

    If dictGrupos.Count = 0 Then
        MsgBox "Nenhum registro valido para processar.", vbExclamation
        Exit Sub
    End If

    '--- Outlook -----------------------------------------------------------
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")

    '--- Criar um rascunho por gestor -------------------------------------
    Dim k As Variant
    For Each k In dictGrupos.Keys

        If k = "(SEM GESTOR IDENTIFICADO)" And Not CRIAR_RASCUNHO_SEM_GESTOR Then
            GoTo ProximoGrupo
        End If

        Set itens = dictGrupos(k)
        Set olMail = olApp.CreateItem(0)                ' 0 = olMailItem

        ' Forca o carregamento da assinatura padrao do Outlook
        Set olInsp = olMail.GetInspector
        assinatura = olMail.HTMLBody

        If k <> "(SEM GESTOR IDENTIFICADO)" Then
            emailDest = dictEmails(k)
            If emailDest <> "" Then
                Set olRcp = olMail.Recipients.Add(emailDest)
            Else
                Set olRcp = olMail.Recipients.Add(dictNomes(k))
            End If
            olRcp.Type = 1                              ' 1 = olTo
            On Error Resume Next
            olRcp.Resolve                               ' tenta resolver na GAL
            On Error GoTo 0
            If CC_FIXO <> "" Then olMail.CC = CC_FIXO
        End If

        olMail.Subject = MontarAssunto(CStr(k) = "(SEM GESTOR IDENTIFICADO)")
        olMail.HTMLBody = MontarCorpoHTML(dictNomes(k), itens, _
                          CStr(k) = "(SEM GESTOR IDENTIFICADO)") & assinatura
        olMail.Save                                     ' salva em Rascunhos

        qtdRascunhos = qtdRascunhos + 1
        Set olMail = Nothing
        Set olInsp = Nothing

ProximoGrupo:
    Next k

    MsgBox "Concluido." & vbCrLf & vbCrLf & _
           "Rascunhos criados: " & qtdRascunhos & vbCrLf & _
           "Colaboradores listados: " & qtdColabs & vbCrLf & _
           "Linhas sem gestor identificado: " & qtdSemGestor & vbCrLf & vbCrLf & _
           "Revise os rascunhos na pasta Rascunhos do Outlook antes do envio.", _
           vbInformation, "Exames Periodicos - Rascunhos"

End Sub

'==============================================================================
' FUNCOES DE APOIO
'==============================================================================

' Mapeia cabecalho (normalizado, sem acentos) -> numero da coluna
Private Function MapearColunas(ws As Worksheet) As Object
    Dim d As Object, c As Long, ultCol As Long, h As String
    Set d = CreateObject("Scripting.Dictionary")
    ultCol = ws.Cells(LINHA_CABECALHO, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultCol
        h = SemAcento(TextoCel(ws.Cells(LINHA_CABECALHO, c).Value))
        If h <> "" And Not d.Exists(h) Then d.Add h, c
    Next c
    Set MapearColunas = d
End Function

Private Function ValidarColunas(d As Object) As Boolean
    Dim obrig As Variant, i As Long, faltando As String
    obrig = Array(H_FUNC, H_MAT, H_GESTOR, H_LOCAL, H_UF, H_CIDADE, _
                  H_EXAME, H_SIT, H_PROX)
    For i = LBound(obrig) To UBound(obrig)
        If Not d.Exists(obrig(i)) Then faltando = faltando & " - " & obrig(i) & vbCrLf
    Next i
    If faltando <> "" Then
        MsgBox "Colunas obrigatorias nao encontradas no cabecalho:" & vbCrLf & _
               faltando & vbCrLf & "Verifique a aba ativa e a linha de cabecalho.", _
               vbCritical, "Estrutura da planilha"
        ValidarColunas = False
    Else
        ValidarColunas = True
    End If
End Function

' Converte valor de celula em texto seguro (trata erro, vazio e numeros)
Private Function TextoCel(ByVal v As Variant) As String
    If IsError(v) Then
        TextoCel = ""
    ElseIf IsEmpty(v) Then
        TextoCel = ""
    ElseIf IsNumeric(v) Then
        If v = Int(v) Then
            TextoCel = Format$(v, "0")
        Else
            TextoCel = CStr(v)
        End If
    Else
        TextoCel = Trim$(CStr(v))
    End If
End Function

' Data em dd/mm/aaaa; aceita data real, serial do Excel, vazio ou erro
Private Function FormatarDataBR(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        FormatarDataBR = "-"
    ElseIf IsDate(v) Then
        FormatarDataBR = Format$(CDate(v), "dd/mm/yyyy")
    ElseIf IsNumeric(v) Then
        If v > 20000 Then
            FormatarDataBR = Format$(CDate(CDbl(v)), "dd/mm/yyyy")
        Else
            FormatarDataBR = "-"
        End If
    Else
        FormatarDataBR = "-"
    End If
End Function

Private Function MontarCidadeUF(ByVal cidade As String, ByVal uf As String) As String
    If cidade <> "" And uf <> "" Then
        MontarCidadeUF = cidade & " - " & uf
    ElseIf cidade <> "" Then
        MontarCidadeUF = cidade
    Else
        MontarCidadeUF = uf
    End If
End Function

' Remove acentos e coloca em minusculas (comparacao robusta de cabecalhos)
Private Function SemAcento(ByVal s As String) As String
    Dim i As Long, cod As Long, ch As String, saida As String
    s = LCase$(Trim$(s))
    For i = 1 To Len(s)
        cod = AscW(Mid$(s, i, 1))
        Select Case cod
            Case 224 To 229: ch = "a"
            Case 231:        ch = "c"
            Case 232 To 235: ch = "e"
            Case 236 To 239: ch = "i"
            Case 241:        ch = "n"
            Case 242 To 246: ch = "o"
            Case 249 To 252: ch = "u"
            Case Else:       ch = Mid$(s, i, 1)
        End Select
        saida = saida & ch
    Next i
    SemAcento = saida
End Function

' Assunto do e-mail (montado com ChrW para preservar acentuacao)
Private Function MontarAssunto(ByVal semGestor As Boolean) As String
    Dim s As String
    s = "A" & ChrW(199) & ChrW(195) & "O NECESS" & ChrW(193) & "RIA | " & _
        "Exames Peri" & ChrW(243) & "dicos da Equipe " & ChrW(8211) & _
        " agenda pela manh" & ChrW(227) & " (at" & ChrW(233) & " 11h) " & _
        ChrW(8211) & " Inspe" & ChrW(231) & ChrW(227) & "o ANVISA"
    If semGestor Then s = "[PENDENTE - SEM GESTOR] " & s
    MontarAssunto = s
End Function

' Corpo do e-mail em HTML (acentos via entidades para nao depender de encoding)
Private Function MontarCorpoHTML(ByVal nomeGestor As String, _
                                 ByVal itens As Collection, _
                                 ByVal semGestor As Boolean) As String
    Dim h As String, item As Variant

    h = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:11pt;color:#333333;max-width:860px;'>"

    ' Barra superior (padrao visual Sonova)
    h = h & "<div style='background:#0083CA;color:#FFFFFF;padding:10px 16px;" & _
            "font-size:12pt;font-weight:bold;'>" & _
            "SA&Uacute;DE OCUPACIONAL | EXAMES PERI&Oacute;DICOS</div>"

    h = h & "<div style='padding:14px 4px 0 4px;'>"

    If semGestor Then
        h = h & "<p style='background:#FFF4E5;border:1px solid #CCCCCC;padding:8px;'>" & _
                "<b>Rascunho interno de pend&ecirc;ncias:</b> os colaboradores abaixo " & _
                "est&atilde;o sem gestor identificado na planilha (c&eacute;lula vazia " & _
                "ou em erro). Ajuste a coluna Gestor e execute novamente, ou trate " & _
                "estes casos manualmente.</p>"
    Else
        h = h & "<p>Prezado(a) <b>" & nomeGestor & "</b>,</p>"
    End If

    h = h & "<p>Estamos conduzindo a regulariza&ccedil;&atilde;o dos " & _
            "<b>exames ocupacionais peri&oacute;dicos</b> pendentes da sua equipe. " & _
            "Os agendamentos ser&atilde;o realizados pelo RH na <b>cl&iacute;nica " & _
            "credenciada mais pr&oacute;xima do local de trabalho</b> de cada " & _
            "colaborador, <b>sempre no per&iacute;odo da manh&atilde;, com " & _
            "atendimento at&eacute; as 11h</b>.</p>"

    h = h & "<p><b>O que pedimos:</b> organize previamente a agenda da equipe para " & _
            "que os colaboradores listados abaixo estejam liberados nesse " & _
            "per&iacute;odo. As convoca&ccedil;&otilde;es individuais, com data, " & _
            "hor&aacute;rio e endere&ccedil;o da cl&iacute;nica, ser&atilde;o " & _
            "enviadas na sequ&ecirc;ncia.</p>"

    h = h & "<p><b>Contexto e prioridade:</b> teremos <b>inspe&ccedil;&atilde;o da " & _
            "ANVISA</b> e a regularidade dos exames peri&oacute;dicos &eacute; " & _
            "<b>ponto de auditoria</b>. O comparecimento dentro do prazo &eacute; " & _
            "indispens&aacute;vel para a conformidade da unidade.</p>"

    ' Tabela de colaboradores
    h = h & "<table style='border-collapse:collapse;width:100%;font-size:10pt;'>"
    h = h & "<tr style='background:#E8F3FA;color:#003C64;'>" & _
            CelCab("Colaborador") & CelCab("Matr&iacute;cula") & _
            CelCab("Local de Trabalho") & CelCab("Cidade / Estado") & _
            CelCab("Exame") & CelCab("Situa&ccedil;&atilde;o") & _
            CelCab("Pr&oacute;x. Vencimento") & "</tr>"

    For Each item In itens
        h = h & "<tr>" & _
                CelDado(CStr(item(0))) & CelDado(CStr(item(1))) & _
                CelDado(CStr(item(2))) & CelDado(CStr(item(3))) & _
                CelDado(CStr(item(4))) & CelDado(CStr(item(5))) & _
                CelDado(CStr(item(6))) & "</tr>"
    Next item
    h = h & "</table>"

    h = h & "<p style='margin-top:14px;'>Em caso de conflito de agenda ou " & _
            "d&uacute;vida, responda este e-mail para buscarmos a melhor " & _
            "alternativa.</p>"

    h = h & "<p>Contamos com o seu apoio.</p>"
    h = h & "<p>Atenciosamente,</p>"
    h = h & "</div></div>"

    MontarCorpoHTML = h
End Function

Private Function CelCab(ByVal txt As String) As String
    CelCab = "<td style='border:1px solid #CCCCCC;padding:6px;font-weight:bold;'>" & _
             txt & "</td>"
End Function

Private Function CelDado(ByVal txt As String) As String
    If txt = "" Then txt = "-"
    CelDado = "<td style='border:1px solid #CCCCCC;padding:6px;'>" & txt & "</td>"
End Function
