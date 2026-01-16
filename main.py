import json
import os
import sys
import questionary
from arquivos import buscar_e_mover_pdf
from planilha import atualizar_status_excel
from log import registrar_log

def carregar_configuracoes():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def iniciar_programa():
    config = carregar_configuracoes()
    
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')

        print("\n" + "="*40)
        print("   SISTEMA DE RECADASTRAMENTO")
        print("="*40)
        
        ano_escolhido = questionary.select(
            "Selecione o ano de referência:",
            choices=["2025", "2026", "SAIR DO SISTEMA"]
        ).ask()
        
        if not ano_escolhido or ano_escolhido == "SAIR DO SISTEMA":
            print("Saindo do sistema... Até logo!")
            break

        nome_pessoa = input("Nome Completo: ").strip().upper()
        
        if not nome_pessoa:
            continue 

        observacao = input("Observação (Enter para SEM ALTERAÇÃO): ").strip()

        caminho_ano = os.path.join(config['caminho_raiz'], f"{config['nome_pasta_ano']} {ano_escolhido}")

        print(f"\n[1/2] 🔍 Buscando PDF de {nome_pessoa}...")
        pasta_origem = buscar_e_mover_pdf(caminho_ano, nome_pessoa, config['subpasta_destino_pdf'])

        if pasta_origem:
            print(f"[2/2] 📝 PDF movido! Agora localizando na planilha...")
            
            sucesso_excel = atualizar_status_excel(
                caminho_ano, 
                pasta_origem, 
                nome_pessoa, 
                observacao
            )
            
            if sucesso_excel:
                print("\n" + "v"*40)
                print("       PROCESSO FINALIZADO COM SUCESSO!")
                print("."*40)

                registrar_log(nome_pessoa, "SUCESSO", f"Ano: {ano_escolhido} | Obs: {observacao}")
        else:
            print("\n" + "!"*40)
            print("       PDF NÃO ENCONTRADO. NADA FOI ALTERADO.")
            print("!"*40)
            
            registrar_log(nome_pessoa, "NAO ENCONTRADO", f"Pasta do ano {ano_escolhido}")

        input("\nPressione [ENTER] para o próximo cadastro...")

if __name__ == "__main__":
    try:
        iniciar_programa()
    except KeyboardInterrupt:
        print("\nPrograma interrompido pelo usuário.")