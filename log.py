from datetime import datetime

def registrar_log(nome, resultado, detalhe=""):
    
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    linha = f"{data_hora} | {nome:<30} | {resultado:<15} | {detalhe}\n"
    
    try:
        with open("historico_execucao.txt", "a", encoding="utf-8") as arquivo:
            arquivo.write(linha)
    except Exception as e:
        print(f"⚠ Erro ao salvar log: {e}")