# Relatório de Vendas por Loja

Este projeto automatiza a geração e envio de um relatório de vendas por loja utilizando Python. O relatório é gerado a partir de um arquivo Excel e enviado por email usando o Microsoft Outlook.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - pandas
  - pywin32

## Instalação

1. Certifique-se de que você tenha o Python instalado em sua máquina. Você pode baixá-lo em [python.org](https://www.python.org/).
2. Instale as bibliotecas necessárias executando os seguintes comandos:
   ```bash
   pip install pandas
   pip install pywin32

1. Coloque o arquivo Vendas.xlsx no mesmo diretório do script Python.
2. Altere o destinatário do email no script (mail.To) para o endereço de email desejado.
3. Execute o script Python:
   python nome_do_seu_script.py

Funcionamento do Script
Leitura do Arquivo Excel:

O script lê o arquivo Vendas.xlsx e armazena os dados na variável tabela_vendas.
Agrupamento e Cálculo de Vendas e Quantidade:

Os dados são agrupados por ID Loja e a soma do Valor Final e da Quantidade são calculadas para cada loja.
Cálculo do Ticket Médio:

O ticket médio é calculado dividindo o Valor Final pela Quantidade para cada loja.
Envio do Email:

Uma instância do Outlook é criada e um novo email é gerado com as tabelas de vendas, quantidade e ticket médio incluídas no corpo do email em formato HTML.
O email é enviado para o destinatário especificado.
