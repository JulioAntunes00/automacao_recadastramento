import os
import shutil

def buscar_e_mover_pdf(caminho_ano, nome_pessoa, subpasta_destino):
    """
    Procura o PDF da pessoa dentro de 'FICHAS EM PDF' e move para 'JA ENVIADOS'.
    """
    pasta_fichas = os.path.join(caminho_ano, "FICHAS EM PDF")

    for nome_pasta_mes in os.listdir(pasta_fichas):
        caminho_da_pasta_atual = os.path.join(pasta_fichas, nome_pasta_mes)

        # Verificamos se é realmente uma pasta (para não tentar abrir arquivos soltos)
        if os.path.isdir(caminho_da_pasta_atual):
            
            # 3. Construímos o caminho que o arquivo teria se estivesse aqui
            caminho_do_pdf = os.path.join(caminho_da_pasta_atual, f"{nome_pessoa}.pdf")

            if os.path.exists(caminho_do_pdf):
                
                # 5. Se existe, verificamos se a pasta de destino (JA ENVIADOS) também existe
                caminho_final_destino = os.path.join(caminho_da_pasta_atual, subpasta_destino)

                if os.path.exists(caminho_final_destino):
                    shutil.move(caminho_do_pdf, os.path.join(caminho_final_destino, f"{nome_pessoa}.pdf"))
                    
                    return nome_pasta_mes
                else:
                    print(f"⚠ Erro: A pasta '{subpasta_destino}' não existe em {nome_pasta_mes}.")
                    return None

    print(f"✖ Erro: O PDF de '{nome_pessoa}' não foi encontrado.")
    return None