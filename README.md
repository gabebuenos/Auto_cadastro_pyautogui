# 🤖 Automação de Cadastro de Produtos via Interface Gráfica

Este script Python foi desenvolvido para automatizar o processo de cadastro de produtos em um sistema, lendo os dados diretamente de uma planilha Excel e simulando ações de usuário (mouse e teclado) para preencher os campos.

---

## ⚙️ Como a Automação Funciona?

O script realiza uma série de ações para extrair informações da planilha e inseri-las na interface gráfica do sistema (RPA - Robotic Process Automation).

### 🛠️ Bibliotecas Principais

O processo é baseado no uso de três bibliotecas Python que trabalham em conjunto:

* **`openpyxl`**: É utilizada para **ler e processar** os dados contidos na planilha `Planilha_Produtoss.xlsx`. O script percorre cada **linha** da planilha, tratando cada linha como um novo produto a ser cadastrado, começando a partir da segunda linha (ignorando o cabeçalho).
* **`pyperclip`**: É a ponte entre a planilha e o sistema. Ela **copia** o valor de cada célula da planilha para a **área de transferência** (clipboard) do sistema operacional.
* **`pyautogui`**: É o motor da automação. Ele **simula ações humanas** (mouse e teclado) para interagir com o sistema:
    * `pyautogui.click(x, y)`: Clica em coordenadas $(x, y)$ específicas na tela para selecionar o campo de preenchimento.
    * `pyautogui.hotkey('ctrl', 'v')`: **Cola** o dado que está na área de transferência no campo selecionado.
    * `pyautogui.scroll(valor)`: Rola a tela (para cima ou para baixo) para que mais campos de cadastro fiquem visíveis.

### 📝 Fluxo Detalhado

O script opera em um loop por cada produto na planilha:

1.  **Extração e Cópia**: Um dado específico (ex: Nome do Produto) é lido da planilha e copiado para a área de transferência usando o `pyperclip`.
2.  **Interação e Colagem**: O `pyautogui.click()` simula um clique na posição exata do campo "Nome do Produto" na tela. Em seguida, `pyautogui.hotkey('ctrl', 'v')` cola o valor copiado.
3.  **Repetição**: Este ciclo se repete para todos os campos do produto (Descrição, Categoria, Preço, etc.) e, depois, para todos os produtos (linhas) da planilha.

### ➕ Controles Específicos

* **Pausas Estratégicas (`time.sleep(3)`)**: O script inclui pausas de 3 segundos para garantir que o sistema de cadastro tenha tempo suficiente para carregar, processar a entrada ou rolar a tela antes que o próximo comando de clique/colagem seja executado.
* **Seleção Condicional (Tamanho)**: Uma lógica `if/elif/else` é implementada para lidar com campos de seleção (como `Pequeno`, `Médio`, `Grande`). O script clica em coordenadas diferentes na tela, dependendo do valor lido na célula da planilha para esse campo.

---

## 🚀 Como Usar a Automação

Para rodar o script com sucesso, você deve preparar seu ambiente e ajustar as configurações específicas da sua tela.

### 1. Instalação de Bibliotecas

Abra seu terminal ou prompt de comando e instale as dependências necessárias:

```bash
pip install openpyxl pyperclip pyautogui
