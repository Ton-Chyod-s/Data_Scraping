# Projetos

### preenchimento.py

Desenvolvi um programa em Python usando várias bibliotecas úteis. O objetivo do programa era percorrer uma planilha e criar um documento Word personalizado com base nos dados encontrados. Para alcançar isso, utilizei as seguintes bibliotecas:

A biblioteca os foi usada para interagir com o sistema operacional, permitindo que eu manipulasse arquivos e pastas necessários para o processo.

A biblioteca docx foi utilizada para criar e manipular documentos Word. Com ela, pude criar, formatar e adicionar conteúdo ao documento Word gerado.

A biblioteca tqdm foi utilizada para exibir uma barra de progresso durante a execução do programa. Isso tornou mais fácil acompanhar o progresso à medida que percorria a planilha e criava o documento Word.

A biblioteca datetime foi usada para obter informações sobre a data e hora atual. Pude usá-la para registrar quando o programa foi executado ou para adicionar informações relevantes ao documento Word.

A biblioteca docx.shared forneceu recursos para definir o tamanho da fonte, o alinhamento do parágrafo e o alinhamento do tabulador no documento Word. Isso me permitiu personalizar a aparência do documento gerado.

A biblioteca openpyxl foi usada para carregar e manipular a planilha de dados. Com ela, pude percorrer as células, obter os valores necessários e criar o documento Word com base nesses dados.

A biblioteca re (expressões regulares) foi utilizada para realizar operações de correspondência de padrões. Com ela, pude fazer verificações e filtragens específicas nos dados obtidos da planilha.

A biblioteca PySimpleGUI foi usada para criar uma interface gráfica do usuário (GUI) para o programa. Ela ofereceu facilidades para criar janelas, botões e outros elementos interativos, tornando a interação com o programa mais amigável.

Ao longo do programa, utilizei estruturas de controle, como def, for, if e try, para iterar sobre os dados da planilha, aplicar lógica condicional, lidar com exceções e executar funções personalizadas quando necessário.
No final, o programa atingiu o resultado desejado. Ele percorreu a planilha, extraiu os dados relevantes, formatou-os em um documento Word personalizado e exibiu uma interface gráfica interativa para facilitar a interação do usuário.

### teste quantidade.py

O código apresentado é um script em Python que realiza algumas operações com arquivos XML e planilhas XLSX.

Aqui está uma descrição do que o código faz:

Importa as bibliotecas necessárias: sleep para adicionar pausas no script, load_workbook do openpyxl para lidar com planilhas XLSX, PySimpleGUI para exibir pop-ups, xml.etree.ElementTree para manipular arquivos XML, random para gerar números aleatórios, shutil para operações de arquivo, e os para operações do sistema.

Verifica se os arquivos necessários, "xxx.xlsx" e "xxx.xlsx", estão presentes na pasta. Caso contrário, exibe uma mensagem de pop-up informando os arquivos ausentes.

Define uma função pasta que cria uma pasta se ela não existir, utilizando o caminho fornecido.

Cria várias pastas usando a função pasta, para organizar os arquivos que serão gerados posteriormente.

Inicia um loop de repetição que itera de 2 a 401 (400 vezes).

Dentro do loop, recupera os valores das células na planilha "xxx.xlsx" e "xxx.xlsx" para as variáveis correspondentes.

Verifica se as coordenadas são nulas. Se forem nulas, o loop é interrompido.

Verifica se o campo de CEP no arquivo "xxx.xlsx" está vazio. Se estiver vazio, o valor de CEP é obtido de outra célula.

Recupera os valores das células nas planilhas "xxxx.xlsx" e "roteiro.xlsx" para outras variáveis correspondentes.

Executa uma sequência de condicionais para determinar o tipo de construção a ser criada: casa, casa com casas secundárias ou prédio.

Para o caso de construção de uma casa, o código lê um arquivo XML de modelo ("xxx.xml"), substitui os valores relevantes e grava um novo arquivo XML. Em seguida, o arquivo XML é compactado em um arquivo zip usando a biblioteca shutil. O arquivo XML original é excluído.

Para o caso de construção de uma casa com casas secundárias, o código lê um arquivo XML de modelo ("xxx.xml") e itera sobre a quantidade de casas secundárias. Para cada casa secundária, o código realiza operações semelhantes às descritas no passo 11.

Para o caso de construção de um prédio, o código lê dois arquivos XML de modelo ("arquivo.xml" e "apartamento.xml"). Ele gera um número aleatório para o nome do arquivo ZIP, substitui os valores relevantes nos arquivos XML e grava os arquivos XML correspondentes. Em seguida, o código movimenta o último arquivo XML gerado para uma pasta específica e realiza outras operações para modificar o XML principal. Por fim, o XML principal é gravado em um arquivo e compactado em um arquivo ZIP. Os arquivos XML originais são excluídos.

Exibe uma mensagem de pop-up informando que a criação foi concluída.

Em resumo, o código realiza a leitura e manipulação de arquivos XML e planilhas XLSX, cria novos arquivos