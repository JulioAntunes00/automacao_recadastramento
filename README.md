# Automação de Recadastramento

Este projeto consiste em um sistema modular desenvolvido em Python para automatizar o fluxo de organização de documentos e atualização de registros de auditoria. O script é responsável por localizar arquivos PDF em diretórios de rede, movê-los para pastas de conclusão e atualizar planilhas de controle em Excel.

---

### Status do Projeto: 🚧**Em Construção**🚧

O projeto encontra-se em fase de desenvolvimento e testes. Atualmente, a estrutura de integração entre os módulos de arquivos e configuração está operacional, com o módulo de manipulação de planilhas em fase de refinamento.

---

### Arquitetura do Sistema

A solução foi projetada seguindo princípios de **Código Limpo** (Clean Code), separando as responsabilidades em arquivos distintos para facilitar a manutenção:

* **main.py**: Atua como o orquestrador do sistema, gerenciando a interface de linha de comando e o fluxo de dados entre os módulos.
* **arquivos.py**: Gerencia operações de sistema de arquivos, incluindo a varredura de diretórios por palavras-chave e a movimentação de documentos.
* **planilha.py**: Responsável pela interação com arquivos .xlsx utilizando a biblioteca openpyxl para busca e edição de células específicas.
* **config.json**: Centraliza parâmetros variáveis como caminhos de rede e nomes de diretórios, permitindo a portabilidade do código.

---

### Fluxo de Funcionamento

1. **Entrada de Dados**: O usuário fornece o nome completo do beneficiário e o ano de referência.
2. **Localização**: O sistema identifica a pasta do mês correspondente dentro do diretório de PDFs através de uma busca por palavra-chave.
3. **Movimentação**: Caso o PDF seja localizado, ele é movido para uma subpasta específica de itens já processados.
4. **Registro**: Após a movimentação do arquivo físico, o script localiza o nome do beneficiário na **Coluna B** da planilha correspondente e registra o status na **Coluna F** e o fechamento na **Coluna H**.

---

### Tecnologias Utilizadas

* **Python 3.13**
* **Bibliotecas Nativas**: os, shutil, json
* **Biblioteca Externa**: openpyxl


---

### Como clonar o projeto no VS Code:

Para levar o código para outro computador, utilize o comando abaixo no terminal:

`git clone https://github.com/JulioAntunes00/automacao_recadastramento .`

Em seguida, instale a dependência necessária:

`pip install openpyxl`