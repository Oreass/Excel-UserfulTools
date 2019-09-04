Attribute VB_Name = "Z_UTEIS"
Option Explicit
Public StartTime As Single

'--- MACROS PADRÕES ---
Sub Z_INI()
    Dim WS As Worksheet
    Dim ATIVA As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ATIVA = ActiveSheet
    For Each WS In ThisWorkbook.Worksheets
        WS.Visible = xlSheetVisible
        WS.Unprotect
    Next
    
    ATIVA.Activate
End Sub

Sub Z_FIM()
    Dim WS As Worksheet
    Dim ATIVA As Worksheet
    
    Set ATIVA = ActiveSheet
    
    GoTo FIM
    
    
    For Each WS In ThisWorkbook.Worksheets
        If Left(WS.Name, 3) = "BD " Or Left(WS.Name, 3) = "OC " Or Left(WS.Name, 3) = "CM " Or Left(WS.Name, 3) = "IN " Or Left(WS.Name, 3) = "TD " Or Left(WS.Name, 3) = "CFG" Then
            WS.Visible = xlSheetHidden
        ElseIf Left(WS.Name, 3) = "ND " Then
            'FAÇA NADA
        Else
            WS.Protect
        End If
    Next
    Application.StatusBar = ""
FIM:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ATIVA.Activate

End Sub

Sub Z_CLASSIFICA_CABECALHO(CABECALHO As Range)
    Dim LINS As Long
    Dim ini As Range
    
    Set ini = CABECALHO.CurrentRegion.Range("A1")
    LINS = ini.CurrentRegion.Rows.Count - 1
    If LINS = 0 Then GoTo FIM
    ini.CurrentRegion.Sort CABECALHO, xlAscending, Header:=xlYes
FIM:
End Sub

Function Z_REMOVER_ACENTO(Caract As String)
 'Referência: https://www.funcaoexcel.com.br/remover-acentos/
 
 Dim A As String
 Dim B As String
 Dim i As Integer
 Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
 Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
 
 For i = 1 To Len(AccChars)
    A = Mid(AccChars, i, 1)
    B = Mid(RegChars, i, 1)
    Caract = Replace(Caract, A, B)
 Next
 
 Z_REMOVER_ACENTO = Caract
End Function

Sub Z_REMOVE_DUPLICATAS(CABECALHO As Range)
    ' CONCIDERA QUE EXISTE CABEÇALHO
    ' REMOVE DUPLICADAS PELA PRIMEIRA LINHA (COLUNA SERIAL)
    '
    
    Dim LINS As Long
    Dim COLS As Long
    Dim IT As Long
    Dim IT2 As Long
    Dim POS As Long
    Dim DADOS() As Variant
    Dim RESULT() As Variant
    
    Dim EXISTE As New Dictionary
    
    COLS = CABECALHO.CurrentRegion.Columns.Count
    LINS = CABECALHO.CurrentRegion.Rows.Count - 1
    If LINS = 0 Then GoTo FIM
    
    ' - CARREGA
    DADOS = CABECALHO.Cells(2, 1).RESIZE(LINS, COLS).Value
    RESULT = DADOS
    
    POS = 0
    ' - MARCA A LINHA
    For IT = 1 To LINS
        If Not IsError(DADOS(IT, 1)) Then
            If Not EXISTE.Exists(UCase(DADOS(IT, 1))) Then
                ' - CRIA
                EXISTE(UCase(DADOS(IT, 1))) = IT
                ' - SALVA
                POS = POS + 1
                For IT2 = 1 To COLS
                    RESULT(POS, IT2) = DADOS(IT, IT2)
                Next IT2
            End If
        End If
    Next IT
    
    ' - DERRUBA
    CABECALHO.CurrentRegion.OFFSET(1, 0).ClearContents
    CABECALHO.OFFSET(1, 0).RESIZE(POS, COLS) = RESULT
    
    Erase DADOS
    Erase RESULT
FIM:
    EXISTE.RemoveAll
End Sub

Function Z_ATUALIZA_COLUNAS_ESPECIFICIAS(ORIGEM As Range, DESTINO As Range, REGRAS As Range, Optional ALIAS As Range = Nothing) As Long
    'PREMISSA: TUDO QUE TENHO NA ORIGEM E É ALTERÁVEL TENHO NO DESTINO
    
    Dim CABE_ORIGEM As New Dictionary
    Dim CABE_DESTINO As New Dictionary
    Dim CABE_ALIAS As New Dictionary
    Dim CHAVES_DESTINO As New Dictionary
    
    Dim DADOS_ORIGEM As Variant
    Dim DADOS_DESTINO As Variant
    Dim DADOS_ALIAS As Variant
    Dim DADOS_REGRAS As Variant
    
    Dim REGRAS_ORIGEM As New Collection 'NÚMERO DO CABEÇALHO NA ORIGEM
    Dim REGRAS_DESTINO As New Collection 'NÚMERO DO CABEÇALHO NO DESTINO
    
    Dim ATUALIZOU As Boolean 'FLAG PARA SABER SE ATUALIZOU
    Dim QTD_ATUA As Long 'QUANTIDADE DE LINHAS ATUALIZADAS
    
    Dim IT As Long
    Dim CHAVE As String
    Dim COLUNA_MD As Long
    Dim POS As Long
    Dim COL As Long
    
    
    
    ' - CABE DE ORIGEM E DESTINO
    CABE_ORIGEM.RemoveAll
    CABE_DESTINO.RemoveAll
    DADOS_ORIGEM = ORIGEM.CurrentRegion
    DADOS_DESTINO = DESTINO.CurrentRegion
    Z_ATUALIZA_COLUNAS_ESPECIFICIAS = 0
    
    For IT = 1 To UBound(DADOS_ORIGEM, 2)
        CABE_ORIGEM(UCase(DADOS_ORIGEM(1, IT))) = IT
    Next IT
    
    For IT = 1 To UBound(DADOS_DESTINO, 2)
        CABE_DESTINO(UCase(DADOS_DESTINO(1, IT))) = IT
    Next IT
    
    ' - CARREGA O ALIAS
    If Not ALIAS Is Nothing Then
        DADOS_ALIAS = ALIAS.CurrentRegion
        For IT = 2 To UBound(DADOS_ALIAS, 1)
            CABE_ALIAS(UCase(DADOS_ALIAS(IT, 2))) = UCase(DADOS_ALIAS(IT, 1)) 'ORIGEM -> DESTINO
        Next IT
        Erase DADOS_ALIAS
    End If
    
    ' - RESETA AS REGRAS
    While REGRAS_ORIGEM.Count > 0
        REGRAS_ORIGEM.Remove (REGRAS_ORIGEM.Count)
    Wend
    While REGRAS_DESTINO.Count > 0
        REGRAS_DESTINO.Remove (REGRAS_DESTINO.Count)
    Wend
    
    ' - DESCOBRE O QUE VAI ATUALIZAR
    DADOS_REGRAS = REGRAS.CurrentRegion.Value2
    For IT = 2 To UBound(DADOS_REGRAS, 1)
        CHAVE = UCase(DADOS_REGRAS(IT, 1))
        ' - VERIFICO SE EXISTE NA BASE
        If CABE_ORIGEM.Exists(CHAVE) Then
            ' - VERIFICA SE POSSUI ALIAS
            If CABE_ALIAS.Exists(CHAVE) Then
                'POSSUI ALIAS
                If CABE_DESTINO.Exists(CABE_ALIAS(CHAVE)) Then
                    REGRAS_ORIGEM.Add CABE_ORIGEM(CHAVE)
                    REGRAS_DESTINO.Add CABE_DESTINO(CABE_ALIAS(CHAVE))
                End If
            Else
                'NÃO POSSUI ALIAS
                If CABE_DESTINO.Exists(CHAVE) Then
                    REGRAS_ORIGEM.Add CABE_ORIGEM(CHAVE)
                    REGRAS_DESTINO.Add CABE_DESTINO(CHAVE)
                End If
            End If
        End If
    Next IT
    
    'TESTA SE TEM ALGO A FAZER
    If REGRAS_ORIGEM.Count = 0 Then GoTo FIM
    
    'VERIFICA SE EXISTE A COLUNA MD NO DESTINO
    COLUNA_MD = 0
    If UCase(DADOS_DESTINO(1, UBound(DADOS_DESTINO, 2))) = "MD" Then COLUNA_MD = UBound(DADOS_DESTINO, 2)
    
    'COLETA OS SERIAIS EXISTENTES
    For IT = 2 To UBound(DADOS_DESTINO, 1)
        CHAVES_DESTINO(UCase(DADOS_DESTINO(IT, 1))) = IT
    Next IT
    
    'FAZ AS ATUALIZAÇÕES
    QTD_ATUA = 0
    For IT = 2 To UBound(DADOS_ORIGEM, 1)
        CHAVE = UCase(DADOS_ORIGEM(IT, 1))
        'VERIFICA SE A CHAVE EXISTE
        If Not CHAVES_DESTINO.Exists(UCase(DADOS_ORIGEM(IT, 1))) Then
            GoTo ERROCHAVE
        Else
            '-SALVA
            POS = CHAVES_DESTINO(CHAVE)
            ATUALIZOU = False
            For COL = 1 To REGRAS_ORIGEM.Count
                If DADOS_ORIGEM(IT, REGRAS_ORIGEM(COL)) <> DADOS_DESTINO(POS, REGRAS_DESTINO(COL)) Then
                    DADOS_DESTINO(POS, REGRAS_DESTINO(COL)) = DADOS_ORIGEM(IT, REGRAS_ORIGEM(COL))
                    ATUALIZOU = True
                End If
            Next COL
            
            ' - FLAGS E CONTAGEM DE ATUALIZAÇÃO
            If ATUALIZOU Then
                If COLUNA_MD <> 0 Then DADOS_DESTINO(POS, COLUNA_MD) = "A"
                QTD_ATUA = QTD_ATUA + 1
            End If
            
        End If
    Next IT
    
    'DERRUBA
    DESTINO.CurrentRegion = DADOS_DESTINO
    Z_ATUALIZA_COLUNAS_ESPECIFICIAS = QTD_ATUA
    
    If False Then
ERROCHAVE:
        MsgBox "Chave " & CHAVE & " não encontrada no DESTINO!", vbCritical, "ExcelExpert"
    End If
FIM:
    
    ' - LIMPA
    Erase DADOS_ORIGEM
    Erase DADOS_DESTINO
    Erase DADOS_ALIAS
    Erase DADOS_REGRAS
    
    CABE_ORIGEM.RemoveAll
    CABE_DESTINO.RemoveAll
    CABE_ALIAS.RemoveAll
    CHAVES_DESTINO.RemoveAll
    
    Set CABE_ORIGEM = Nothing
    Set CABE_DESTINO = Nothing
    Set CABE_ALIAS = Nothing
    Set CHAVES_DESTINO = Nothing
    
    While REGRAS_ORIGEM.Count > 0
        REGRAS_ORIGEM.Remove (REGRAS_ORIGEM.Count)
    Wend
    While REGRAS_DESTINO.Count > 0
        REGRAS_DESTINO.Remove (REGRAS_DESTINO.Count)
    Wend
    Set REGRAS_ORIGEM = Nothing
    Set REGRAS_DESTINO = Nothing
    
End Function

Function Z_ATUALIZA(ORIGEM As Range, DESTINO As Range, Optional VERIFICA_DELETE As Boolean = False, Optional ALIAS As Range = Nothing) As String
    ' - IMPORTA OS DADOS DE UM LOCAL PARA OUTRO VERIFICANDO DUPLICIDADE
    ' - SE A ÚLTIMA COLUNA SE CHAMA "MD" ALÉM DE ATUALIZAR/INSERIR COLOCO "A"
    ' - SE VERIFICA_DELETE = true ENTÃO REMOVE DA BASE OU FLAGA A COLUNA "MD" COM "D"
    '
    ' - PRIMEIRA COLUNA DE AMBOS OBRIGATÓRIO O SERIAL !
    ' - VERIFICA DELETE INSERE O "D" NO DESTINO OU EXCLUI CASO NÃO EXISTA A COLUNA "MD"
    '
    
    Dim DADOS_OR() As Variant
    Dim DADOS_DEST() As Variant
    Dim NOVOS() As Variant
    
    Dim ORI As New Dictionary
    Dim DES As New Dictionary
    
    Dim IT As Long
    Dim COL As Long
    Dim NOVO As Long
    
    Dim TXT As String 'APOIO A VERIFICAÇÃO DE DUPLICIDADES
    Dim DUP As String 'APOIO A VERIFICAÇÃO DE DUPLICIDADES
    
    Dim COLUNA_MD As Integer 'COLUNA PARA ATUALIZAÇÃO EXTERNA
    Dim PERMANECE As Long 'CONTA O NÚMERO DE LINHAS QUE PERMANECEM NA BASE
    
    Dim DELETADOS As Long 'REMOÇÃO DE DADOS
    Dim ATUALIZADOS As Long 'CONTA OS ITENS ATUALIZADOS DE FATO
    
    COLUNA_MD = 0
    DELETADOS = 0
    ATUALIZADOS = 0
    
    ' - CARREGA OS DADOS
    DADOS_OR = ORIGEM.CurrentRegion.Value
    DADOS_DEST = DESTINO.CurrentRegion.Value
    If DADOS_DEST(1, UBound(DADOS_DEST, 2)) = "MD" Then COLUNA_MD = UBound(DADOS_DEST, 2)
    ReDim NOVOS(1 To UBound(DADOS_OR, 1), 1 To UBound(DADOS_DEST, 2))
    
    ' - CARREGA OS CABECALHOS DO DESTINO
    DUP = ""
    If Not ALIAS Is Nothing Then
        'COM ALIAS
        Dim DADOS As Variant
        Dim DIC_ALIAS As New Dictionary
        
        DADOS = ALIAS.CurrentRegion.Value2
        For IT = 2 To UBound(DADOS, 1)
            DIC_ALIAS(UCase(DADOS(IT, 2))) = DADOS(IT, 1) 'DESCOBRE O NOME CORRETO PELO ALIAS (PESQUISA->BANCO)
        Next IT
        
        For IT = 1 To UBound(DADOS_DEST, 2)
            TXT = ""
            If DIC_ALIAS.Exists(UCase(DADOS_DEST(1, IT))) Then
                'COM ALIAS DEFINIDO
                TXT = DIC_ALIAS(UCase(DADOS_DEST(1, IT)))
            Else
                'SEM ALIAS DEFINIDO
                TXT = DADOS_DEST(1, IT)
            End If
            
            ' - TESTA SE O CABEÇALHO ESTÁ DUPLICADO
            If DES.Exists(TXT) Then
                If DUP <> "" Then DUP = DUP & ","
                DUP = DUP & TXT & "(" & IT & ")"
            End If
            
            DES(TXT) = IT
        Next IT
        
        Erase DADOS
        DIC_ALIAS.RemoveAll
        Set DIC_ALIAS = Nothing
    Else
        'SEM ALIAS
        For IT = 1 To UBound(DADOS_DEST, 2)
            TXT = DADOS_DEST(1, IT)
            
            If DES.Exists(TXT) Then
                If DUP <> "" Then DUP = DUP & ", "
                DUP = DUP & TXT & "(" & IT & ")"
            End If
            
            DES(TXT) = IT
        Next IT
    End If
    
    If DUP <> "" Then
        MsgBox "Cabeçalhos [" & DUP & "] duplicados!"
        Err.Raise 1993, , "Ajuste os cabeçalhos antes de prosseguir"
    End If
    
    ' - ASSOCIA O CABEÇALHO DE ORIGEM COM DESTINO
    'SEM ALIAS
    ORI.RemoveAll
    For IT = 1 To UBound(DADOS_OR, 2)
        If DES.Exists(DADOS_OR(1, IT)) Then _
            ORI(DADOS_OR(1, IT)) = DES(DADOS_OR(1, IT))
    Next IT
    
    ' - CARREGA A POSIÇÃO DE CADA SERIAL (PRIMEIRA COLUNA)
    DES.RemoveAll
    For IT = 2 To UBound(DADOS_DEST, 1)
        DES(CStr(DADOS_DEST(IT, 1))) = IT
    Next IT
    
    ' - CARREGA
    NOVO = 0
    For IT = 2 To UBound(DADOS_OR, 1)
        If DES.Exists(CStr(DADOS_OR(IT, 1))) Then
            ' - EXISTE
            If DES(CStr(DADOS_OR(IT, 1))) <> "NOVO" Then
                ' - CASO NORMAL
                If COLUNA_MD = 0 Then
                    For COL = 1 To UBound(DADOS_OR, 2)
                        If ORI.Exists(DADOS_OR(1, COL)) Then
                            If Len(DADOS_OR(IT, COL)) = 10 And IsDate(DADOS_OR(IT, COL)) Then DADOS_OR(IT, COL) = CDate(DADOS_OR(IT, COL))
                            DADOS_DEST(DES(CStr(DADOS_OR(IT, 1))), ORI(DADOS_OR(1, COL))) = DADOS_OR(IT, COL)
                            ATUALIZADOS = ATUALIZADOS + 1
                        End If
                    Next COL
                Else
                ' - CASO ATUALIZAÇÃO EXTERNA (EXISTE A COLUNA MD)
                    For COL = 1 To UBound(DADOS_OR, 2)
                        If ORI.Exists(DADOS_OR(1, COL)) Then
                            ' - TESTA IGUALDADE
                            If DADOS_DEST(DES(CStr(DADOS_OR(IT, 1))), ORI(DADOS_OR(1, COL))) <> DADOS_OR(IT, COL) Then
                                If Len(DADOS_OR(IT, COL)) = 10 And IsDate(DADOS_OR(IT, COL)) Then DADOS_OR(IT, COL) = CDate(DADOS_OR(IT, COL))
                                DADOS_DEST(DES(CStr(DADOS_OR(IT, 1))), ORI(DADOS_OR(1, COL))) = DADOS_OR(IT, COL)
                                DADOS_DEST(DES(CStr(DADOS_OR(IT, 1))), COLUNA_MD) = "A"
                                ATUALIZADOS = ATUALIZADOS + 1
                            End If
                        End If
                    Next COL
                End If
                
            End If
        Else
            ' - NOVO
            DES(CStr(DADOS_OR(IT, 1))) = "NOVO"
            NOVO = NOVO + 1
            For COL = 1 To UBound(DADOS_OR, 2)
                If ORI.Exists(DADOS_OR(1, COL)) Then
                    If Len(DADOS_OR(IT, COL)) = 10 And IsDate(DADOS_OR(IT, COL)) Then DADOS_OR(IT, COL) = CDate(DADOS_OR(IT, COL))
                    NOVOS(NOVO, ORI(DADOS_OR(1, COL))) = DADOS_OR(IT, COL)
                End If
            Next COL
            
            ' - PARA ATUALIZAÇÃO EXTERNA
            If COLUNA_MD <> 0 Then NOVOS(NOVO, COLUNA_MD) = "A"
            
        End If
    Next IT
    
    ' - NÚMERO DE LINHAS ANTIGAS QUE PERMANECEM NA BASE
    PERMANECE = UBound(DADOS_DEST, 1)
    
    ' - COLOCA A FLAG "D" OU REMOVE A LINHA
    If VERIFICA_DELETE = True Then
        Dim SER_ATUAIS As New Dictionary 'ARMAZENA OS SERIAIS ATUAIS
        SER_ATUAIS.RemoveAll
        
        ' - COLETA OS SERIAIS DA ORIGEM
        For IT = 2 To UBound(DADOS_OR)
            SER_ATUAIS(CStr(DADOS_OR(IT, 1))) = IT
        Next IT
        
        ' - REMOÇÃO
        If COLUNA_MD <> 0 Then
            ' - FLAGA MD
            For IT = 2 To UBound(DADOS_DEST, 1)
                If Not SER_ATUAIS.Exists(CStr(DADOS_DEST(IT, 1))) Then
                    DELETADOS = DELETADOS + 1
                    DADOS_DEST(IT, COLUNA_MD) = "D"
                End If
            Next IT
        Else
            ' - DELETA
            PERMANECE = 1
            For IT = 2 To UBound(DADOS_DEST, 1)
                If Not SER_ATUAIS.Exists(CStr(DADOS_DEST(IT, 1))) Then
                    DELETADOS = DELETADOS + 1
                    DADOS_DEST(IT, COLUNA_MD) = "D"
                Else
                    PERMANECE = PERMANECE + 1
                    For COL = 1 To UBound(DADOS_DEST, 2)
                        DADOS_DEST(PERMANECE, COL) = DADOS_DEST(IT, COL)
                    Next COL
                End If
            Next IT
        End If
        
        ' - LIMPA
        SER_ATUAIS.RemoveAll
        Set SER_ATUAIS = Nothing
    End If
    
    ' - DERRUBA
    DESTINO.CurrentRegion.OFFSET(1, 0).ClearContents
    DESTINO.RESIZE(PERMANECE, UBound(DADOS_DEST, 2)) = DADOS_DEST
    
    If NOVO > 0 Then _
        DESTINO.Cells(PERMANECE + 1, 1).RESIZE(NOVO, UBound(DADOS_DEST, 2)) = NOVOS
    
    ' - RETORNO
    Z_ATUALIZA = "Operação realizada com sucesso!"
    If ATUALIZADOS > 0 Then Z_ATUALIZA = Z_ATUALIZA & vbNewLine & "Atualizados:" & ATUALIZADOS & " registro(s)."
    If NOVO > 0 Then Z_ATUALIZA = Z_ATUALIZA & vbNewLine & "Novos:" & NOVO & " registro(s)."
    If DELETADOS > 0 Then Z_ATUALIZA = Z_ATUALIZA & vbNewLine & "Removidos:" & DELETADOS & " registro(s)."
                 
    ' - LIMPA A MEMÓRIA
    Erase DADOS_DEST
    Erase DADOS_OR
    Erase NOVOS
    ORI.RemoveAll
    DES.RemoveAll
    Set ORI = Nothing
    Set DES = Nothing
    
End Function

Function Z_IMPORTA(ORIGEM As Range, DESTINO As Range) As Long
    ' - IMPORTA OS DADOS DE UM LOCAL PARA O OUTOR SEM VERIFICAR DUPLICIDADE

    Dim DADOS_OR() As Variant
    Dim DADOS_DEST() As Variant
    Dim NOVOS() As Variant
    
    Dim ORI As New Dictionary
    Dim DES As New Dictionary
    
    Dim IT As Long
    Dim COL As Long
    Dim NOVO As Long
    
    
    ' - CARREGA OS DADOS
    DADOS_OR = ORIGEM.CurrentRegion.Value
    DADOS_DEST = DESTINO.CurrentRegion.RESIZE(2).Value
    
    ReDim NOVOS(1 To UBound(DADOS_OR, 1), 1 To UBound(DADOS_DEST, 2))
    
    ' - CARREGA OS CABECALHOS DO DESTINO
    For IT = 1 To UBound(DADOS_DEST, 2)
        DES(DADOS_DEST(1, IT)) = IT
    Next IT
    
    ' - ASSOCIA O CABEÇALHO DE ORIGEM COM DESTINO
    ORI.RemoveAll
    For IT = 1 To UBound(DADOS_OR, 2)
        If DES.Exists(DADOS_OR(1, IT)) Then _
            ORI(DADOS_OR(1, IT)) = DES(DADOS_OR(1, IT))
    Next IT
    
    ' - CARREGA
    NOVO = 0
    For IT = 2 To UBound(DADOS_OR, 1)
        ' - NOVO
        NOVO = NOVO + 1
        For COL = 1 To UBound(DADOS_OR, 2)
            If ORI.Exists(DADOS_OR(1, COL)) Then
                If Len(DADOS_OR(IT, COL)) = 10 And IsDate(DADOS_OR(IT, COL)) Then DADOS_OR(IT, COL) = CDate(DADOS_OR(IT, COL))
                NOVOS(NOVO, ORI(DADOS_OR(1, COL))) = DADOS_OR(IT, COL)
            End If
        Next COL
    Next IT
    
    ' - DERRUBA
    IT = DESTINO.CurrentRegion.Rows.Count + 1
    If NOVO > 0 Then _
        DESTINO.Cells(IT, 1).RESIZE(NOVO, UBound(DADOS_DEST, 2)) = NOVOS
    
    ' - RETORNO
    Z_IMPORTA = NOVO
    
    ' - LIMPA A MEMÓRIA
    Erase DADOS_DEST
    Erase DADOS_OR
    Erase NOVOS
    ORI.RemoveAll
    DES.RemoveAll
    Set ORI = Nothing
    Set DES = Nothing
    
End Function
Sub Z_OCULTAR(PLANILHA As Worksheet, CHAVES As Range, Optional REFAZ As Boolean = False)
    ' - COLUNA 1: DESCRIÇÃO
    ' - COLUNA 2: FLAG
    ' - COLUNA 3: ANTIGA
    ' - COLUNA 4: COLUNA DE
    ' - COLUNA 5: COLUNA PARA
    
    Dim IT As Long
    Dim TAM As Long
    
    IT = 2
    While CHAVES.Cells(IT, 1) <> ""
        If CHAVES.Cells(IT, 2) <> CHAVES.Cells(IT, 3) Or REFAZ Then
            TAM = CHAVES.Cells(IT, 5) - CHAVES.Cells(IT, 4) + 1
            
            ' - OCULTA
            If CHAVES.Cells(IT, 2) = True Then
                PLANILHA.Cells(1, CHAVES.Cells(IT, 4)).RESIZE(1, TAM).EntireColumn.Hidden = False
            Else
                PLANILHA.Cells(1, CHAVES.Cells(IT, 4)).RESIZE(1, TAM).EntireColumn.Hidden = True
            End If
            
            ' - SALVA
            CHAVES.Cells(IT, 3) = CHAVES.Cells(IT, 2)
            
        End If
        IT = IT + 1
    Wend
    
End Sub

Sub Z_FILTRAR(ORIGEM As Range, DESTINO As Range, CRIT As Range, COLUNAS As Integer)
    ' - REMOVE OS FILTROS
    ORIGEM.CurrentRegion.AutoFilter
    ORIGEM.CurrentRegion.AutoFilter
    
    ' - LIMPA O DESTINO
    DESTINO.CurrentRegion.OFFSET(1, 0).Clear
    ORIGEM.CurrentRegion.AdvancedFilter xlFilterCopy, CRIT, DESTINO.RESIZE(1, COLUNAS)
    
    ' - INSERE AS FÓRMULAS
    Call Z_INSERE_FORMULAS(DESTINO, True)
    Call Z_INSERE_FORMULAS(DESTINO, False, -3)
    
    ' - FORMATA
    Dim LISN As Long
    Dim COLS As Long
    LINS = Range("CRUZ_EMITIDO").CurrentRegion.Rows.Count - 1
    COLS = Range("CRUZ_EMITIDO").CurrentRegion.Columns.Count
    
    Range("CRUZ_EMITIDO").Cells(-1, 1).RESIZE(1, COLS).Copy
    Range("CRUZ_EMITIDO").Cells(2, 1).RESIZE(LINS, COLS).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Sub Z_REMOVER_LINHAS(INI_DADOS As Range, CHAVES, COLUNA As Integer, Optional LIMPAR_TUDO As Boolean = True)
    Dim DADOS() As Variant
    Dim RESULT() As Variant
    Dim CHAVE() As Variant
    Dim OK As Boolean
    
    Dim LINS As Long
    Dim COLS As Long
    
    Dim POS As Long
    Dim IT As Long
    Dim IT2 As Long
    Dim ITC As Long
    
    ' - TESTA SE AS CHAVES SÃO ARRAYS
    If IsArray(CHAVES) Then
        CHAVE = CHAVES
    Else
        ReDim CHAVE(1 To 1, 1 To 1)
        CHAVE(1, 1) = CHAVES
    End If
    
    ' - PROCESSO
    LINS = INI_DADOS.CurrentRegion.Rows.Count - 1
    COLS = INI_DADOS.CurrentRegion.Columns.Count
    If LINS < 1 Then GoTo FIM
    
    DADOS = INI_DADOS.CurrentRegion.Value
    RESULT = DADOS
    POS = 0
    
    For IT = 2 To UBound(DADOS, 1)
        ' - COPIE
        OK = True
        
        ' - TESTA SE NÃO É NENHUMA DAS CHAVES
        For IT2 = 1 To UBound(CHAVE, 1)
            If DADOS(IT, COLUNA) = CHAVE(IT2, 1) Then
                OK = False
            End If
        Next IT2
        
        'SALVA O RESULTADO
        If OK Then
            POS = POS + 1
            For ITC = 1 To UBound(DADOS, 2)
                RESULT(POS, ITC) = DADOS(IT, ITC)
            Next ITC
        End If
    Next IT
    
    If LIMPAR_TUDO Then
        INI_DADOS.CurrentRegion.OFFSET(1, 0).Clear 'LIMPA TUDO
    Else
        INI_DADOS.CurrentRegion.OFFSET(1, 0).ClearContents 'LIMPA APENAS O CONTEÚDO
    End If
    
    If POS > 0 Then
        INI_DADOS.CurrentRegion.Range("A1").Cells(2, 1).RESIZE(POS, COLS) = RESULT
    End If
    
    Erase DADOS
    Erase RESULT
    
FIM:
End Sub

Sub Z_INSERE_FORMULAS(INI_CABECALHO As Range, Optional REMOVER As Boolean = True, Optional OFF_FORMULA As Integer = -2, Optional POS As Long = 2, Optional TAMANHO = 0)
    Dim LINS As Long
    Dim COLS As Long
    
    
    LINS = INI_CABECALHO.CurrentRegion.Rows.Count - 1
    If TAMANHO <> 0 Then LINS = TAMANHO
    
    
    COLS = INI_CABECALHO.CurrentRegion.Columns.Count
    If LINS < 1 Then GoTo FIM
    
    Application.Calculation = xlCalculationManual
    
    INI_CABECALHO.RESIZE(1, COLS).OFFSET(OFF_FORMULA, 0).Copy
    INI_CABECALHO.Cells(POS, 1).RESIZE(LINS, COLS).PasteSpecial xlPasteFormulas, , True
    
    Application.Calculate
    
    If REMOVER Then
        INI_CABECALHO.Cells(POS, 1).RESIZE(LINS, COLS).Copy
        INI_CABECALHO.Cells(POS, 1).RESIZE(LINS, COLS).PasteSpecial xlPasteValues
    End If
    
    Application.CutCopyMode = False
    
    Application.Calculation = xlCalculationAutomatic
    
FIM:
End Sub

Sub Z_INSERE_FORMATOS(INI_CABECALHO As Range, Optional CEL_COL_FORMULA As Integer = -1)
    Dim LINS As Long
    Dim COLS As Long
    LINS = INI_CABECALHO.CurrentRegion.Rows.Count - 1
    COLS = INI_CABECALHO.CurrentRegion.Columns.Count
    If LINS > 0 Then
        INI_CABECALHO.Cells(CEL_COL_FORMULA, 1).RESIZE(1, COLS).Copy
        INI_CABECALHO.Cells(2, 1).RESIZE(LINS, COLS).PasteSpecial xlPasteFormats
        INI_CABECALHO.Cells(2, 1).RESIZE(LINS, COLS).PasteSpecial xlPasteValidation
        Application.CutCopyMode = False
    End If
End Sub


Function Z_OBTERARQUIVO(Optional descFILTRO As String = "Todos os arquivos", Optional extFILTRO As String = "*.*")
    Dim fDlg As FileDialog
    Dim RESULT As String
    
    'Chama o objeto passando os parâmetros
    Set fDlg = Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)
    
    RESULT = ""
    
    With fDlg
        'Alterar esta propriedade para True permitirá a seleção de vários arquivos
        .AllowMultiSelect = False
 
        'Determina a forma de visualização dos aruqivos
        .InitialView = msoFileDialogViewDetails
        
        'Lipa os filtros
        .Filters.Clear
        
        'Filtro de arquivos, pode ser colocado mais do que um filtro separando com ; por exemplo: "*.xls;*.xlsm"
        .Filters.Add descFILTRO, extFILTRO
        
        'Determina qual o drive inicial
        .InitialFileName = ""
    End With
 
    'Retorna o arquivo selecionado
    If fDlg.Show = -1 Then
        RESULT = fDlg.SelectedItems(1)
    Else
        RESULT = ""
    End If
    
    'SALVA O CAMINHO
    Z_OBTERARQUIVO = RESULT
End Function

Function Z_OBTERARPASTA()
    Dim fDlg As FileDialog
    Dim RESULT As String
    
    'Chama o objeto passando os parâmetros
    Set fDlg = Application.FileDialog(FileDialogType:=msoFileDialogFolderPicker)
    
    RESULT = ""
    
    With fDlg
        'Alterar esta propriedade para True permitirá a seleção de vários arquivos
        .AllowMultiSelect = False
 
        'Determina a forma de visualização dos aruqivos
        .InitialView = msoFileDialogViewDetails
        
        'Lipa os filtros
        .Filters.Clear
        
        'Filtro de arquivos, pode ser colocado mais do que um filtro separando com ; por exemplo: "*.xls;*.xlsm"
        '.Filters.Add descFILTRO, extFILTRO
        
        'Determina qual o drive inicial
        .InitialFileName = ""
    End With
 
    'Retorna o arquivo selecionado
    If fDlg.Show = -1 Then
        RESULT = fDlg.SelectedItems(1)
    Else
        RESULT = ""
    End If
    
    'SALVA O CAMINHO
    Z_OBTERARQUIVO = RESULT
End Function

Sub Z_AJUSTAR_COLUNA(INI_COLUNA As Range, FORMULA As String)
    'Altera os valores da coluna conforme regra
    
    Dim NUMLINS As Long
    
    NUMLINS = INI_COLUNA.CurrentRegion.Rows.Count - 1
    
    If NUMLINS = 0 Then GoTo FIM
    
    '1 - Insere uma coluna auxiliar
    INI_COLUNA.OFFSET(0, 1).EntireColumn.Insert
    
    '2 - Insere a fórmula
    INI_COLUNA.OFFSET(1, 1).RESIZE(NUMLINS, 1).FormulaR1C1 = FORMULA
    
    '3 - Atualiza os valores
    INI_COLUNA.OFFSET(1, 1).RESIZE(NUMLINS, 1).Copy
    INI_COLUNA.OFFSET(1, 0).RESIZE(NUMLINS, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    '4 - Remove a coluna auxiliar
    INI_COLUNA.OFFSET(1, 1).EntireColumn.Delete
FIM:
End Sub

Sub Z_ATUALIZA_TABELAS_DINAMICAS()
    ' ATUALIZA TODAS AS TABELAS DINÂMICAS, INDEPENDE DE A_xx
    Dim WS As Worksheet
    Set WS = ActiveSheet
    ActiveWorkbook.RefreshAll
    WS.Activate
End Sub

 Sub Z_TIMER_START()
'   anota o valor do timer no open do arquivo ( ou a qualquer momento...)
    StartTime = Timer
End Sub

Sub Z_TIMER_SHOW()
    Dim X As Double
    Dim EndTime As Single
    Dim PastTime As Single
    Const Conversor = 86400 'constante: 86.400 segundos em um dia

    EndTime = Timer     'anota o valor do time
    PastTime = EndTime - StartTime  ' calcula o tempo decorrido
    PastTime = PastTime / Conversor ' converte em horas
    MsgBox "Tempo de Uso   " & Format(PastTime, "Long Time")

End Sub

Sub Z_ATUALISTAS()
    Dim POS As Long
    Dim LINS As Long
    Dim COLS As Long
    
    Call Z_INI
    
    POS = 0
    While Range("INI_LST").OFFSET(0, POS) <> ""
        ' - LINHAS E COLUNAS DA TABELA
        LINS = Range("INI_LST").OFFSET(0, POS).CurrentRegion.Rows.Count - 1
        LINS = WorksheetFunction.MAX(LINS, 1)
        COLS = Range("INI_LST").OFFSET(0, POS).CurrentRegion.Columns.Count
        
        ' - TABELA OU LISTA
        If Range("INI_LST").OFFSET(0, POS).CurrentRegion.Columns.Count = 1 Then
            Range("INI_LST").OFFSET(1, POS).RESIZE(LINS, COLS).Name = "LST_" & Range("INI_LST").OFFSET(0, POS)
        Else
            Range("INI_LST").OFFSET(1, POS).RESIZE(LINS, COLS).Name = "TAB_" & Range("INI_LST").OFFSET(0, POS)
            Range("INI_LST").OFFSET(1, POS).RESIZE(LINS, 1).Name = "LST_" & Range("INI_LST").OFFSET(0, POS)
        End If
        
        ' - PRÓXIMA LISTA/TABELA
        POS = POS + Range("INI_LST").OFFSET(0, POS).CurrentRegion.Columns.Count + 1
    Wend
    
    Call Z_FIM
End Sub

Sub Z_ATUALIZA_ARQUIVOS(PASTA As String)
    Dim IT As Long
    Dim ARQ As Long
    
    Dim BUSCA As Range
    Dim FSO As New FileSystemObject
    Dim ARQUIVO As File
    Dim PREFIXO As String
    Dim PR(1 To 9) As Variant
    
    Dim STATUS_ARQ As Boolean
    
    ' "CFG_INI_ATUA_EXT"
    '
    ' - VARRE UMA TABELA NO FORMATO ESPECÍFICO E INSERE OS DADOS
    'FORMATO:
    'PREFIXO DO ARQUIVO  PLANILHA    CÉLULA  INFORMAÇÃO  TABELA EXT  LIMPAR  L   C   OBSERVAÇÃO
    
    If Not FSO.FolderExists(ThisWorkbook.Path & "\" & PASTA) Then
        MsgBox "PASTA [" & PASTA & "] INEXISTENTE!", vbCritical, "ERRO"
        GoTo FIM
    End If
    
    IT = 1
    PREFIXO = ""
    STATUS_ARQ = False
    
    Application.DisplayAlerts = False
    
    While Range("CFG_INI_ATUA_EXT").OFFSET(IT, 0) <> ""
        ' PROCURA O ARQUIVO
        If PREFIXO <> Range("CFG_INI_ATUA_EXT").OFFSET(IT, 0) Then
            ' - FECHA O ARQUIVO (SE FOI ABERTO)
            If PREFIXO <> "" And STATUS_ARQ Then
                MsgBox "ARQUIVO ATUALIZADO COM SUCESSO [" & ARQUIVO.Name & "].", vbOKOnly, "ATENÇÃO"
                Workbooks(ARQUIVO.Name).Close True
            End If
            
            ' - PROCURA O ARQUIVO
            PREFIXO = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 0)
            For Each ARQUIVO In FSO.GetFolder(ThisWorkbook.Path & "\" & PASTA).Files
                If UCase(Left(ARQUIVO.Name, Len(PREFIXO))) = UCase(PREFIXO) Then
                    GoTo OK
                End If
            Next
            ' - NÃO ACHOU O ARQUIVO
            STATUS_ARQ = False
            
            ' - ACHOU O ARQUIVO
            If False Then
OK:
                Workbooks.Open ARQUIVO.Path
                ThisWorkbook.Activate
                STATUS_ARQ = True
            End If
        End If
        
        ' SE ACHOU O ARQUIVO FAZ O PROCEDIMENTO
        If STATUS_ARQ Then
            ' CARREGA OS PARÂMETROS
            PR(1) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 1) 'PLANILHA
            PR(2) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 2) 'CELULA
            PR(3) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 3) 'INFO
            PR(4) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 4) 'TABELA
            PR(5) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 5) 'LIMPAR
            PR(6) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 6) 'L
            PR(7) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 7) 'C
            PR(8) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 8) 'OF L
            PR(9) = Range("CFG_INI_ATUA_EXT").OFFSET(IT, 9) 'OF C
            
            ' LIMPA OS DADOS FIXOS
            If PR(5) = "SIM" Then
                ' - DELIMITADO OU TUDO
                If PR(6) = "" Or PR(7) = "" Then
                    ' - TUDO
                    Workbooks(ARQUIVO.Name).Sheets(CStr(PR(1))).Range(CStr(PR(2))). _
                    CurrentRegion.OFFSET(CLng(PR(8)), CLng(PR(9))).ClearContents
                Else
                    ' - DELIMITADO
                    Workbooks(ARQUIVO.Name).Sheets(CStr(PR(1))).Range(CStr(PR(2))). _
                    RESIZE(CLng(PR(6)), CLng(PR(7))).OFFSET(CLng(PR(8)), CLng(PR(9))).ClearContents
                End If
            End If
            
            ' INSERE OS DADOS
            If PR(3) <> "" Or PR(4) = "" Then
                Workbooks(ARQUIVO.Name).Sheets(CStr(PR(1))).Range(CStr(PR(2))). _
                OFFSET(CLng(PR(8)), CLng(PR(9))) = PR(3)
            Else
                Range(CStr(PR(4))).Copy
                Workbooks(ARQUIVO.Name).Sheets(CStr(PR(1))).Range(CStr(PR(2))). _
                OFFSET(CLng(PR(8)), CLng(PR(9))).PasteSpecial xlPasteValues
                Application.CutCopyMode = False
            End If
        End If
        
        IT = IT + 1
    Wend
    
FIM:
    If PREFIXO <> "" And STATUS_ARQ Then
        MsgBox "ARQUIVO ATUALIZADO COM SUCESSO [" & ARQUIVO.Name & "].", vbOKOnly, "ATENÇÃO"
        Workbooks(ARQUIVO.Name).Close True
    End If
    Application.DisplayAlerts = True
End Sub

Function Z_ENCONTRE_CELULAS(VALOR As Variant, INTERVALO As Range) As Range
Dim C As Range, FoundCells As Range
Dim firstaddress As String

Set FoundCells = Nothing

With INTERVALO
    'find first cell that contains "rec"
    Set C = .Find(VALOR, LookIn:=xlValues, LOOKAT:=xlWhole)
    
    'if the search returns a cell
    If Not C Is Nothing Then
        'note the address of first cell found
        firstaddress = C.Address
        Do
            'FoundCells is the variable that will refer to all of the
            'cells that are returned in the search
            If FoundCells Is Nothing Then
                Set FoundCells = C
            Else
                Set FoundCells = Union(C, FoundCells)
            End If
            'find the next instance of "rec"
            Set C = .FindNext(C)
        Loop While Not C Is Nothing And firstaddress <> C.Address
    End If
End With

Set Z_ENCONTRE_CELULAS = FoundCells
End Function

Function Z_ARQUIVO_ESTA_ABERTO(FileName As String) As Boolean
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    Z_ARQUIVO_ESTA_ABERTO = False
    Case 70:   Z_ARQUIVO_ESTA_ABERTO = True
    Case Else: Z_ARQUIVO_ESTA_ABERTO = False
    End Select
End Function

Function Z_ESTA_EM_TABELA(CELULA As Range) As Boolean
    On Error GoTo ERRO
    Z_ESTA_EM_TABELA = CELULA.ListObject.Name <> ""
    
    If False Then
ERRO:
        Z_ESTA_EM_TABELA = False
        Err.Clear
    End If
End Function

' -------------------------- MACROS DE INTERFACE

Sub z_Display_Reexibir()
'Attribute Display_Reexibir.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
'   reexibe os titulos
    ActiveWindow.DisplayHeadings = True
'   reexibe as barra de formulas
    Application.DisplayFormulaBar = True
'   reexibe a faixa de opcoes
    Application.DisplayFullScreen = False
'    Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",true)"
End Sub
Sub z_Display_Ocultar()
'
    Application.ScreenUpdating = False

'   oculta a faixa de opcoes
    Application.DisplayFullScreen = True
'   oculta os titulos
    ActiveWindow.DisplayHeadings = False
'   oculta as barra de f—rmulas
    Application.DisplayFormulaBar = False
'   Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",false)"
End Sub
