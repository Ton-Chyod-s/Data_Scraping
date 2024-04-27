# Automatização de Geração e Preenchimento Automático de Documentos em Python.

Este projeto Python consiste em um script que engloba tanto o preenchimento automático de documentos do Word a partir de dados de uma planilha do Excel quanto a manipulação de arquivos XML para geração de informações específicas

## Descrição
O script é capaz de:

* Ler dados de planilhas do Excel.
* Preencher automaticamente documentos do Word com esses dados.
* Manipular e gerar arquivos XML para criar informações específicas.
* Organizar arquivos em diretórios conforme necessário para o processamento dos dados.

## Funcionalidades
* Utilização da biblioteca **openpyxl** para leitura de planilhas do Excel.
* Preenchimento de documentos do Word com a biblioteca **python-docx**.
* Manipulação de arquivos XML para geração de informações personalizadas.
* Criação e organização de diretórios para armazenamento dos arquivos gerados.

## Requisitos
* Python 3.x
* Bibliotecas Python necessárias:
    * **openpyxl**
    * **python-docx**
    * **xml.etree.ElementTree**
    * **PySimpleGUI**

## Como Usar
1. Instale as bibliotecas necessárias:
    ```BASH
    pip install openpyxl python-docx
2. Modifique o código para se adequar aos seus requisitos específicos, como nomes de arquivos, caminhos e lógica de processamento dos dados.
3. Execute o script Python:
    ```BASH
    python seu_script.py

## Observações
Certifique-se de ter o Microsoft Word instalado para manipulação dos documentos do Word.
Os arquivos XML e Excel devem estar presentes na pasta de trabalho do script ou especificados corretamente no código.