# Excel-UserfulTools
Funções em excel responsáveis por automatizar tarefas empresariais comuns.

Sub Z_INI()

Sub Z_FIM()

Sub Z_CLASSIFICA_CABECALHO(CABECALHO As Range)

Function Z_REMOVER_ACENTO(Caract As String)

Sub Z_REMOVE_DUPLICATAS(CABECALHO As Range)

Function Z_ATUALIZA_COLUNAS_ESPECIFICIAS(ORIGEM As Range, DESTINO As Range, REGRAS As Range, Optional ALIAS As Range = Nothing) As Long
    'PREMISSA: TUDO QUE TENHO NA ORIGEM E É ALTERÁVEL TENHO NO DESTINO

Function Z_ATUALIZA(ORIGEM As Range, DESTINO As Range, Optional VERIFICA_DELETE As Boolean = False, Optional ALIAS As Range = Nothing) As String
    ' - IMPORTA OS DADOS DE UM LOCAL PARA OUTRO VERIFICANDO DUPLICIDADE
    ' - SE A ÚLTIMA COLUNA SE CHAMA "MD" ALÉM DE ATUALIZAR/INSERIR COLOCO "A"
    ' - SE VERIFICA_DELETE = true ENTÃO REMOVE DA BASE OU FLAGA A COLUNA "MD" COM "D"
    '
    ' - PRIMEIRA COLUNA DE AMBOS OBRIGATÓRIO O SERIAL !
    ' - VERIFICA DELETE INSERE O "D" NO DESTINO OU EXCLUI CASO NÃO EXISTA A COLUNA "MD"

Function Z_IMPORTA(ORIGEM As Range, DESTINO As Range) As Long
    ' - IMPORTA OS DADOS DE UM LOCAL PARA O OUTOR SEM VERIFICAR DUPLICIDADE

Sub Z_OCULTAR(PLANILHA As Worksheet, CHAVES As Range, Optional REFAZ As Boolean = False)
    ' - COLUNA 1: DESCRIÇÃO
    ' - COLUNA 2: FLAG
    ' - COLUNA 3: ANTIGA
    ' - COLUNA 4: COLUNA DE
    ' - COLUNA 5: COLUNA PARA
    ' - refaz = true para refazer todo o estudo de ocultar
    
Sub Z_FILTRAR(ORIGEM As Range, DESTINO As Range, CRIT As Range, COLUNAS As Integer)

Sub Z_REMOVER_LINHAS(INI_DADOS As Range, CHAVES, COLUNA As Integer, Optional LIMPAR_TUDO As Boolean = True)

Sub Z_INSERE_FORMULAS(INI_CABECALHO As Range, Optional REMOVER As Boolean = True, Optional OFF_FORMULA As Integer = -2, Optional POS As Long = 2, Optional TAMANHO = 0)

Sub Z_INSERE_FORMATOS(INI_CABECALHO As Range, Optional CEL_COL_FORMULA As Integer = -1)

Function Z_OBTERARQUIVO(Optional descFILTRO As String = "Todos os arquivos", Optional extFILTRO As String = "*.*")

Function Z_OBTERARPASTA()

Sub Z_AJUSTAR_COLUNA(INI_COLUNA As Range, FORMULA As String)
    'Altera os valores da coluna conforme formula   R1C1
    
Sub Z_ATUALIZA_TABELAS_DINAMICAS()
    ' ATUALIZA TODAS AS TABELAS DINÂMICAS, INDEPENDE DE A_xx
    
 Sub Z_TIMER_START()
'   anota o valor do timer no open do arquivo ( ou a qualquer momento...)

Sub Z_TIMER_SHOW()

Sub Z_ATUALISTAS()
    ' Varre as células da esquerda para a direita dando o nome para os dados de TAB_ (Primeiro cabeçalho) ou LST_ (primeiro cabeçalho)

Sub Z_ATUALIZA_ARQUIVOS(PASTA As String)
    ' "CFG_INI_ATUA_EXT"
    '
    ' - VARRE UMA TABELA NO FORMATO ESPECÍFICO E INSERE OS DADOS
    'FORMATO:
    'PREFIXO DO ARQUIVO  PLANILHA    CÉLULA  INFORMAÇÃO  TABELA EXT  LIMPAR  L   C   OBSERVAÇÃO

Function Z_ENCONTRE_CELULAS(VALOR As Variant, INTERVALO As Range) As Range
    'DADO UMA SELEÇÃO QUALQUER (NÃO PRECISA SER CONTÍNUA), RETORNA OUTRA SELEÇÃO CONTENDO O VALOR INDICADO

Function Z_ARQUIVO_ESTA_ABERTO(FileName As String) As Boolean
    'TESTA SE UM ARQUIVO ESTÁ ABERTO (QUALQUER)

Function Z_ESTA_EM_TABELA(CELULA As Range) As Boolean
    'TESTA SE A CÉLULA SELECIONADA ESTÁ EM UMA TABELA

' -------------------------- MACROS DE INTERFACE

Sub z_Display_Reexibir()
    'EXIBE O EXCEL

Sub z_Display_Ocultar()
    'OCULTA O EXCEL
