import json
import os
import questionary
from arquivos import buscar_e_mover_pdf
from planilha import atualizar_status_excel

def carregar_configuracoes():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def iniciar_programa():
    config = carregar_configuracoes()
    
    print("\n=== SISTEMA DE RECADASTRAMENTO ===")
    
    ano_escolhido = questionary.select(
        "Selecione o ano de referência:",
        choices=["2025", "2026"]
    ).ask()
    
    if not ano_escolhido:
        return

    nome_pessoa = input("Nome Completo: ").strip().upper()
    observacao = input("Observação (Enter para 'SEM ALTERAÇÃO'): ").strip()

    caminho_ano = os.path.join(
        config['caminho_raiz'], 
        f"{config['nome_pasta_ano']} {ano_escolhido}"
    )

    if not os.path.exists(caminho_ano):
        print(f"✖ Erro: Pasta do ano {ano_escolhido} não encontrada.")
        return

    print("------------------------------------------------")
    print("1. Procurando PDF...")
    pasta_origem_encontrada = buscar_e_mover_pdf(
        caminho_ano, 
        nome_pessoa, 
        config['subpasta_destino_pdf']
    )

    if pasta_origem_encontrada:
        print("------------------------------------------------")
        print("2. Atualizando Planilha...")
        
        atualizar_status_excel(
            caminho_ano, 
            pasta_origem_encontrada,
            nome_pessoa, 
            observacao
        )
    
    print("------------------------------------------------")
    print("Processo Finalizado.\n")

if __name__ == "__main__":
    iniciar_programa()