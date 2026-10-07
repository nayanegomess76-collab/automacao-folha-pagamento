Automação de Folha de Pagamento

Automação desenvolvida em Python para extrair informações de folhas de pagamento em PDF e atualizar automaticamente uma planilha Excel.

Objetivo

O projeto automatiza uma tarefa operacional de processamento de folha de pagamento, reduzindo a necessidade de lançamento manual de informações no Excel.

A automação realiza a leitura de um PDF com os recibos de pagamento, identifica os colaboradores e extrai os principais valores necessários para atualização da planilha.

O que a automação faz
Lê recibos de pagamento em PDF
Identifica o colaborador em cada página
Extrai o valor líquido do pagamento
Extrai os valores de FGTS
Extrai os valores de INSS
Normaliza os nomes para facilitar a correspondência entre PDF e Excel
Atualiza automaticamente a planilha de folha de pagamento
Atualiza a planilha de tributos
Mantém os demais dados existentes na planilha
Ignora colaboradores que não estejam na planilha correspondente
Realiza uma verificação após o salvamento do arquivo
Fluxo do processo

PDF de folha de pagamento → Extração dos dados → Tratamento e limpeza → Correspondência dos colaboradores → Atualização do Excel → Validação

Tecnologias utilizadas
Python
pdfplumber
openpyxl
Regex
Excel
Dados extraídos
Informação	Destino
Valor líquido	Planilha de folha de pagamento
FGTS	Planilha de tributos
INSS	Planilha de tributos
Validação

Após a atualização, o arquivo Excel é salvo e reaberto para verificar se os valores foram efetivamente gravados.

Essa etapa aumenta a confiabilidade do processo e facilita a identificação de possíveis falhas durante a execução.


Privacidade dos dados

Este repositório contém somente o código da automação.

Arquivos reais de folha de pagamento, planilhas, PDFs, nomes de colaboradores, salários e demais informações pessoais ou financeiras não fazem parte do projeto público.

Para testes, devem ser utilizados dados fictícios ou anonimizados.


Instalação

Instale as dependências do projeto:
pip install -r requirements.txt
Configuração
kkkkk
Antes de executar a automação, informe os caminhos dos arquivos PDF e Excel nas variáveis correspondentes do script.


Execução

python automacao_folha.py


Objetivo profissional

Este projeto demonstra a aplicação de Python em uma situação real de automação de processos administrativos e financeiros, envolvendo:
Extração de dados
Tratamento e padronização
Automação de tarefas repetitivas
Integração entre PDF e Excel
Validação de dados
Redução de trabalho manual

