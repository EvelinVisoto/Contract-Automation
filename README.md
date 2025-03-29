# Automação de Contratos

## Descrição
Este projeto automatiza a geração de contratos utilizando Python. O sistema lê dados de uma planilha Excel e preenche um modelo de documento Word (.docx), mantendo a formatação original. Ele também gera valores numéricos por extenso, valida e formata campos como CPF, CNPJ, endereços e datas, garantindo consistência e precisão na geração dos documentos.

## Funcionalidades
- Leitura de dados a partir de um arquivo Excel (.xlsx)
- Preenchimento automatizado de contratos no formato Word (.docx)
- Manutenção da formatação original do documento
- Geração de valores numéricos por extenso
- Validação e formatação de CPF, CNPJ, endereços e datas
- Geração automatizada e organização dos contratos em uma pasta de saída

## Tecnologias Utilizadas
- **Python 3.x**: Linguagem principal do projeto
- **pandas**: Para manipulação e leitura dos dados do Excel
- **python-docx**: Para manipulação de documentos Word
- **num2words**: Para conversão de números em texto por extenso
- **os**: Para gerenciar diretórios e arquivos

## Como Usar
1. Prepare um arquivo Excel (`dados_funcionarios.xlsx`) com os campos necessários.
2. Insira um modelo de contrato no formato Word (`modelo_contrato.docx`) com as tags correspondentes.
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
│   └── contratos/           # Pasta onde os contratos gerados serão armazenados
```

## Autor
Desenvolvido por **Evelin Visoto** - 2024/2025  
📌 **Repositório no GitHub**: [https://github.com/EvelinVisoto/automacao-de-contratos](https://github.com/EvelinVisoto/automacao-de-contratos)
