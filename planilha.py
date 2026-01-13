import os
from openpyxl import load_workbook

def atualizar_status_excel(caminho_ano, nome_pasta_pdf, nome_pessoa, observacao):
    """
    1. Descobre qual o mês baseado na pasta do PDF.
    2. Encontra a planilha correspondente a esse mês.
    3. Procura a pessoa na Coluna B e preenche as Colunas F e H.
    """
    
    # --- PARTE 1: Encontrar o Arquivo Excel Correto ---
    pasta_excel_raiz = os.path.join(caminho_ano, "PLANILHA DE CONTROLE")
    
    # Vamos descobrir qual mês estamos tratando (ex: "JANEIRO")
    # A pasta do PDF é algo como "A - JANEIRO PLANO 10". Vamos usar isso para achar o Excel.
    arquivo_excel_encontrado = None
    
    # Lista todos os arquivos na pasta de planilhas
    if not os.path.exists(pasta_excel_raiz):
        print("✖ Erro: Pasta 'PLANILHA DE CONTROLE' não encontrada.")
        return False

    arquivos_existentes = os.listdir(pasta_excel_raiz)

    # Lógica: Se a pasta do PDF tem "JANEIRO", queremos o Excel que tem "JANEIRO"
    meses_possiveis = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", 
                       "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
    
    mes_identificado = ""
    for mes in meses_possiveis:
        if mes in nome_pasta_pdf.upper():
            mes_identificado = mes
            break
            
    if not mes_identificado:
        print(f"✖ Erro: Não consegui identificar o mês na pasta '{nome_pasta_pdf}'.")
        return False

    # Agora buscamos o arquivo Excel que contém esse mês no nome
    for arquivo in arquivos_existentes:
        if mes_identificado in arquivo.upper() and arquivo.endswith(".xlsx"):
            arquivo_excel_encontrado = arquivo
            break
    
    if not arquivo_excel_encontrado:
        print(f"✖ Erro: Nenhuma planilha encontrada contendo '{mes_identificado}'.")
        return False

    caminho_completo_excel = os.path.join(pasta_excel_raiz, arquivo_excel_encontrado)
    print(f"📂 Abrindo planilha: {arquivo_excel_encontrado}...")

    # --- PARTE 2: Editar o Excel ---
    try:
        # Carrega o arquivo (data_only=False para manter fórmulas se houver, mas aqui só lemos texto)
        wb = load_workbook(caminho_completo_excel)
        
        pessoa_encontrada = False

        # Procura em TODAS as abas (Plano 10, Plano 11, etc.)
        for nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            
            # Vamos percorrer linha por linha
            for linha in ws.iter_rows(min_row=2): # Começa na linha 2 (pula cabeçalho)
                # Coluna B é o índice 1 na lista da linha (0=A, 1=B, 2=C...) 
                # OU acessamos direto pela célula
                celula_nome = linha[1] # Coluna B
                
                # Verificamos se a célula não está vazia e se é a pessoa
                if celula_nome.value and str(celula_nome.value).strip().upper() == nome_pessoa:
                    
                    # Achamos! Agora editamos.
                    # Coluna F (Status) -> linha[5]
                    # Coluna H (OK) -> linha[7]
                    
                    # Lógica do Status: Se o usuário não digitou nada, é "SEM ALTERAÇÃO"
                    texto_obs = observacao.strip() if observacao else "VIA WHATS SEM ALTERAÇÃO"
                    
                    linha[5].value = texto_obs  # Escreve na Coluna F
                    linha[7].value = "OK"       # Escreve na Coluna H
                    
                    pessoa_encontrada = True
                    print(f"✓ Atualizado na aba '{nome_aba}': {texto_obs}")
                    break # Para de procurar linhas nesta aba
            
            if pessoa_encontrada:
                break # Se achou numa aba, não precisa procurar nas outras

        if pessoa_encontrada:
            wb.save(caminho_completo_excel)
            print("💾 Planilha salva com sucesso!")
            return True
        else:
            print(f"⚠ Aviso: '{nome_pessoa}' não foi encontrado(a) na planilha.")
            return False

    except PermissionError:
        print("✖ ERRO CRÍTICO: A planilha está aberta no Excel!")
        print("➜ Por favor, FECHE o arquivo Excel e tente novamente.")
        return False
    except Exception as e:
        print(f"✖ Erro inesperado ao mexer no Excel: {e}")
        return False