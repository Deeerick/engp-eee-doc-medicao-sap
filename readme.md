<!-- 1 - Ler arquivo excel
2 - Abrir SAP
3 - Acessar transacao IW41
4 - Iterar sobre cada linha do excel
5 - Colocar o número da ordem de manutencao e confirmar
6 - Se exibir uma mensagem, clicar em SIM
7 - Verificar se as colunas "Inicio trabalho" e "Fim trabalho" estao preenchidas
8 - Se a coluna "Inicio trabalho" estiver preenchida, selecionar linha com operacao 0030 no SAP
9 - Se a coluna "Fim trabalho" estiver preenchida, selecionar linha com operacao 0040 no SAP
10 - Clicar no botao "Dados reais" (F8)
11 - Se exibir uma mensagem, clicar em SIM
12 - Preencher as informacoes (Trabalho real, Inicio trabalho, Fim trabalho e Txt confirmacao)
13 - Clicar no botao "Documentos de medicao"
14 - Se exibir uma mensagem de alerta, teclar enter (2x se necessário)
15 - Preencher as informacoes (Data e hora de medicao e lido por) -->


## Descrição

Este projeto automatiza o processo de confirmação de ordens de manutenção no SAP utilizando um arquivo Excel como fonte de dados.

## Requisitos

- SAP GUI instalado
- Python
- Bibliotecas Python: requirements.txt
