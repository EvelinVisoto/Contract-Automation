# Automacao de Contratos

## Descricao
Este projeto automatiza a geracao de contratos utilizando Python. O sistema le dados de uma planilha Excel e preenche um modelo de documento Word (.docx), mantendo a formatacao original. Ele tambem gera valores numericos por extenso, valida e formata campos como CPF, CNPJ, enderecos e datas.

## Funcionalidades
- Leitura de dados a partir de um arquivo Excel (.xlsx)
- Preenchimento automatizado de contratos no formato Word (.docx)
- Manutencao da formatacao original do documento
- Geracao de valores numericos por extenso
- Validacao e formatacao de CPF, CNPJ, enderecos e datas
- Geracao automatizada e organizacao dos contratos em uma pasta de saida

## Tecnologias Utilizadas
- **Python 3.x**: Linguagem principal do projeto
- **pandas**: Para manipulacao e leitura dos dados do Excel
- **python-docx**: Para manipulacao de documentos Word
- **num2words**: Para conversao de numeros em texto por extenso
- **os**: Para gerenciar diretorios e arquivos

## Instalacao
1. Clone este repositório:
   ```sh
   git clone https://github.com/EvelinVisoto/automacao-de-contratos.git
   ```
2. Acesse o diretório do projeto:
   ```sh
   cd automacao-de-contratos
   ```
3. Instale as dependências necessárias:
   ```sh
   pip install pandas python-docx num2words
   ```

## Como Usar
1. Prepare um arquivo Excel (dados_funcionarios.xlsx) com os campos necessários.
2. Insira um modelo de contrato no formato Word (modelo_contrato.docx) com as tags correspondentes.
3. Execute o script:
   ```sh
   python gerar_contratos.py
   ```
4. Os contratos preenchidos serão gerados na pasta `dist/contratos/`.

## Estrutura do Projeto
```
automacao-de-contratos/
│── dados_funcionarios.xlsx  # Planilha com os dados
│── modelo_contrato.docx     # Modelo do contrato
│── gerar_contratos.py       # Script principal
│── dist/
│   └── contratos/           # Pasta onde os contratos gerados serao armazenados
```

## Contribuicao
Sinta-se à vontade para contribuir! Basta fazer um fork, criar uma branch com suas alteracoes e abrir um pull request.

## Autor
Desenvolvido por **Evelin Visoto C. Fernandes** - 2024/2025

## Licenca
Este projeto é de código aberto e está licenciado sob a MIT License.

