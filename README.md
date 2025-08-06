# Auto_cadastro
Esta automação em Python foi criada para agilizar o processo de cadastro de produtos em um sistema ou site, transferindo dados de uma planilha do Excel diretamente para campos na tela.

Esta automação em Python foi criada para agilizar o processo de cadastro de produtos em um sistema ou site, transferindo dados de uma planilha do Excel diretamente para campos na tela.

Como a automação funciona?
O script utiliza três bibliotecas principais:

openpyxl: Lê os dados de uma planilha chamada Planilha_Produtoss.xlsx. O código percorre cada linha da planilha, a partir da segunda linha (a primeira linha é geralmente o cabeçalho).

pyperclip: Copia o valor de cada célula da planilha para a área de transferência.

pyautogui: Simula ações do mouse e do teclado para interagir com a interface de um sistema.

pyautogui.click(): Clica em coordenadas específicas na tela para selecionar campos de preenchimento.

pyautogui.hotkey('ctrl', 'v'): Cola o valor que foi copiado para a área de transferência no campo selecionado.

pyautogui.scroll(): Rola a tela para cima ou para baixo para que mais campos fiquem visíveis.

O script extrai um dado (como o nome do produto) da planilha, copia esse dado, clica na posição exata onde o campo "nome do produto" está na tela e cola a informação. Ele repete esse processo para todos os campos, como descrição, categoria, preço, etc., e para cada linha (produto) da planilha.

Existem também alguns controles específicos:

time.sleep(3): Pausa o script por 3 segundos para dar tempo para o sistema carregar ou para a tela rolar, garantindo que o próximo comando seja executado corretamente.

Seleção de Tamanho: O script inclui uma lógica condicional (if/elif/else) para selecionar o tamanho do produto (Pequeno, Médio, Grande) clicando em posições diferentes, dependendo do valor na planilha. Isso é útil para menus de seleção.

Como usar a automação
Para usar esta automação, você precisará preparar seu ambiente:

Instale as bibliotecas necessárias:

Bash

pip install openpyxl pyperclip pyautogui
Organize sua planilha:

Crie uma planilha no Excel com o nome Planilha_Produtoss.xlsx.

Renomeie a aba (sheet) para Planilha1.

Organize as colunas na mesma ordem em que o script espera os dados:

Coluna A: Nome do produto

Coluna B: Descrição

Coluna C: Categoria

... e assim por diante. A ordem é crucial.

Ajuste as coordenadas:

As coordenadas (1225,352, por exemplo) são específicas para a resolução da sua tela e a posição exata da janela do sistema de cadastro. Você precisará ajustá-las para que o mouse clique nos lugares certos no seu computador.

Para descobrir as coordenadas, use o próprio pyautogui. Abra o terminal, digite python para entrar no interpretador e execute:

Python

import pyautogui
pyautogui.displayMousePosition()
Mova o mouse para os campos que você quer preencher e anote as coordenadas X e Y. Depois, substitua os valores no código.

Execute o script:

Abra o sistema de cadastro na primeira tela de preenchimento.

Execute o script Python. Importante: Não mexa no mouse nem no teclado enquanto o script estiver rodando para não interromper a automação.
