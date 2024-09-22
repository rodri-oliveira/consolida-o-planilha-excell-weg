# Projeto de Consolidação e Filtragem de Planilhas

Este projeto foi desenvolvido para consolidar múltiplas planilhas Excel em um único arquivo e aplicar filtros a partir de critérios definidos pelo usuário. Ele possui uma interface gráfica (GUI) construída com `Tkinter`, possibilitando uma interação simples e amigável para consolidação e geração de relatórios.

## Funcionalidades

- Consolidação de várias planilhas Excel em um único arquivo.
- Geração de relatórios filtrados com base em critérios definidos pelo usuário.
- Exportação de planilhas e relatórios para pastas dedicadas.

## Estrutura do Projeto

```bash
|-- app.py
|-- consolidar.py
|-- main.py
|-- planilhas/              # Pasta contendo 3 planilhas para testes
|-- planilhas-consolidadas/  # Pasta para o arquivo consolidado
|-- relatórios/              # Pasta onde os relatórios gerados pelos filtros serão salvos

Pré-requisitos
Certifique-se de ter os seguintes requisitos instalados:

Python 3.7+
pip (gerenciador de pacotes do Python)
Como baixar o projeto
Para clonar este repositório, use o comando:

bash
Copiar código
git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
Instalação das Dependências
Todas as dependências estão listadas no arquivo requirements.txt. Para instalar, execute o seguinte comando:

bash
Copiar código
pip install -r requirements.txt
Se o arquivo requirements.txt não estiver disponível, você pode instalar as dependências manualmente:

bash
Copiar código
pip install pandas tkinter openpyxl
Estrutura de Pastas
O projeto requer que você tenha duas pastas pré-configuradas:

Pasta de planilhas para teste: Deve conter planilhas Excel (.xlsx) que serão consolidadas.

Local: C:\consolidar-planilha-weg\planilhas
Pasta para arquivos consolidados: O arquivo consolidado será salvo nesta pasta.

Local: C:\consolidar-planilha-weg\relatórios
Pasta para relatórios gerados pelos filtros: Os relatórios gerados pelos filtros aplicados serão salvos nesta pasta.

Local: C:\consolidar-planilha-weg\relatórios
Certifique-se de que essas pastas estejam criadas antes de executar o projeto.

Como Executar o Projeto
Para iniciar o projeto, execute o arquivo main.py:

bash
Copiar código
python main.py
Isso abrirá a interface gráfica para que você possa selecionar o diretório das planilhas, consolidar os dados e aplicar os filtros desejados.

Selecione a pasta das planilhas: Clique no botão "Buscar Pasta" e escolha a pasta onde suas planilhas estão armazenadas.
Consolidar Planilhas: Após selecionar a pasta, clique em "Consolidar". O arquivo consolidado será salvo em C:\consolidar-planilha-weg\planilhas-consolidadas\planilha_consolidada.xlsx.
Gerar Relatórios Filtrados:
Escolha uma coluna do DataFrame.
Selecione um operador (maior que, menor que, ou igual a).
Insira um valor numérico para o filtro.
Clique em "Gerar Relatório". O relatório será salvo na pasta C:\consolidar-planilha-weg\relatórios.
