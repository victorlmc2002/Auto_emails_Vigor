# Auto_emails_Vigor

Automatize o envio de boletos por e-mail para inquilinos e proprietários da empresa Vigor - Gestão de Negócios Imobiliários.

## Objetivo

Este projeto facilita o envio automático de boletos e relatórios financeiros para inquilinos e proprietários, utilizando uma interface gráfica simples e a planilha já utilizada pela empresa.

## Pré-requisitos

- Python 3 instalado na máquina
- Biblioteca `openpyxl` instalada  
  Instale com:
  ```sh
  pip install openpyxl
  ```

## Como usar

1. Certifique-se de que a planilha `BoletosFormatados.xlsx` está no diretório correto (`C:/Users/victo/Desktop/boletos/`).
2. Coloque os arquivos PDF dos boletos, taxas de condomínio e repasses nas pastas correspondentes, conforme o mês vigente.
3. Execute o arquivo principal:
   ```sh
   python app.py
   ```
4. Utilize a interface gráfica para selecionar os imóveis e enviar os e-mails.

## Estrutura do Projeto

- [`app.py`](app.py): Script principal com a interface gráfica e lógica de envio.
- [`devedores.py`](devedores.py): Classe para manipulação dos dados dos devedores.
- `lg.png`: Imagem utilizada nos e-mails.
- `README.md`: Este arquivo.

## Observações

- O sistema utiliza a planilha já existente da empresa, sem necessidade de adaptações.
- Os e-mails são enviados utilizando as credenciais configuradas no código.

---