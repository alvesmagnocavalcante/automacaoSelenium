# Script de Automação para Inserção de Dados de Produtos

Este script Python automatiza o processo de inserção de dados de produtos em uma aplicação web utilizando Selenium. O script lê os dados de um arquivo Excel e preenche os campos correspondentes no site.

## Instruções

1. **Requisitos**
   - Python 3.x
   - Selenium
   - OpenPyXL
   - PyAutoGUI

2. **Instalação**
   - Instale os pacotes necessários utilizando o pip:
     ```
     pip install selenium openpyxl pyautogui
     ```

3. **Utilização**
   - Coloque seus dados de produtos em um arquivo Excel chamado `dados.xlsx`. Certifique-se de que os dados estejam organizados nas colunas A a Q representando diferentes atributos.

   - Execute o script utilizando o Python:
     ```
     python insercao_dados_produto.py
     ```

4. **WebDriver**
   - O script utiliza o WebDriver do Chrome. Certifique-se de tê-lo instalado e o caminho configurado corretamente.

5. **Observação**
   - Ajuste os XPaths no script para que correspondam à estrutura do seu formulário web, se necessário.

6. **Aviso Legal**
   - Utilize este script de forma responsável e em conformidade com os termos de uso do site alvo, que foi feito por DevAprender para fins didáticos. 
