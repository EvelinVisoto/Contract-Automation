# Contract Automation

### Description
This project automates contract generation using Python. The system reads data from an Excel spreadsheet and fills in a Word document template (.docx), maintaining the original formatting. It also converts numerical values to text, validates and formats fields such as CPF, CNPJ, addresses, and dates, ensuring consistency and accuracy in document generation.

### Features
- Reads data from an Excel (.xlsx) file  
- Automatically fills contracts in Word (.docx) format  
- Maintains the original document formatting  
- Converts numerical values to text  
- Validates and formats CPF, CNPJ, addresses, and dates  
- Automatically generates and organizes contracts in an output folder  

### Technologies Used
- **Python**: Main language of the project  
- **pandas**: For handling and reading Excel data  
- **python-docx**: For Word document manipulation  
- **num2words**: For converting numbers into text  
- **os**: For managing directories and files  

### How to Use
1. Prepare an Excel file (`dados_funcionarios.xlsx`) with the required fields.  
2. Insert a contract template in Word format (`modelo_contrato.docx`) with the corresponding tags.  
3. Run the script:  
   ```sh
   python preenchimento_contratos.py
   ```
4. The filled contracts will be generated in the `dist/contratos/` folder.  

### Project Structure
```
Projeto_Contratos/
│── dados_funcionarios.xlsx  # Spreadsheet with data
│── modelo_contrato.docx     # Contract template
│── gerar_contratos.py       # Main script
│── dist/
│   └── contratos/           # Folder where generated contracts will be stored
```

### Author
Developed by **Evelin Visoto** - 2024/2025  
📌 **GitHub Repository**: [https://github.com/EvelinVisoto/Contract-Automation](https://github.com/EvelinVisoto/Contract-Automation) 
