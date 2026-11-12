
# Automação WhatsApp com Python

Projeto acadêmico desenvolvido durante o curso Técnico em Desenvolvimento de Sistemas da ETEC, com o objetivo de explorar automação de tarefas utilizando Python.

A aplicação realiza a leitura de números de telefone armazenados em uma planilha Excel e utiliza o WhatsApp Web para demonstrar a automação de interações com contatos.

O projeto foi desenvolvido como uma experiência prática envolvendo manipulação de arquivos, processamento de dados e automação de interfaces gráficas.

## Funcionalidades

- **Leitura de planilhas Excel:** utilização de arquivos `.xlsx` para obter números de telefone.
- **Processamento de dados:** leitura e verificação de condições específicas nos números armazenados.
- **Integração com WhatsApp Web:** abertura de conversas por meio de links personalizados.
- **Automação de interface:** utilização do PyAutoGUI para simular interações com o navegador.
- **Execução automatizada:** processamento sequencial dos registros da planilha.
- **Compatibilidade com executável:** identificação automática do caminho dos arquivos quando executado como script Python ou executável.

## Tecnologias utilizadas

- **Python:** linguagem principal do projeto.
- **OpenPyXL:** leitura e manipulação de planilhas Excel.
- **PyAutoGUI:** automação de interações com teclado e interface gráfica.
- **Webbrowser:** abertura de páginas no navegador.
- **URL Encode:** formatação de mensagens para utilização em URLs.

## Bibliotecas utilizadas

| Biblioteca | Finalidade |
|---|---|
| `sys` | Identificação do ambiente de execução |
| `openpyxl` | Leitura de arquivos Excel |
| `urllib.parse` | Codificação de mensagens para URLs |
| `webbrowser` | Abertura do navegador |
| `time` | Controle de intervalos de execução |
| `pyautogui` | Automação de teclado |
| `os` | Manipulação de caminhos e arquivos |

## Estrutura do projeto

```text
Teste_Spam_Whatsapp/
│
├── main.py
├── numeros.xlsx
└── README.md
```

- `main.py`: código principal da automação.
- `numeros.xlsx`: planilha contendo os números utilizados durante a execução.
- `README.md`: documentação do projeto.

## Como executar

### Pré-requisitos

- Python instalado.
- Google Chrome ou outro navegador compatível.
- WhatsApp Web autenticado no navegador.
- Arquivo Excel com os números de telefone.

### 1. Clone o repositório

```bash
git clone https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git
```

### 2. Entre na pasta do projeto

```bash
cd Teste_Spam_Whatsapp
```

### 3. Instale as dependências

```bash
pip install openpyxl pyautogui
```

### 4. Configure a planilha

Coloque o arquivo `numeros.xlsx` na mesma pasta do script.

A planilha deve possuir uma aba chamada `Plan1`, contendo os números de telefone na primeira coluna.

### 5. Execute o projeto

```bash
python main.py
```

## Funcionamento

O funcionamento do projeto é dividido em algumas etapas:

1. Identificação do diretório onde o programa está sendo executado.
2. Verificação da existência da planilha Excel.
3. Leitura dos números armazenados na primeira coluna.
4. Verificação de uma condição específica nos números.
5. Abertura do WhatsApp Web por meio do navegador.
6. Automação de interações com a interface utilizando PyAutoGUI.

## Contexto acadêmico

Este projeto foi desenvolvido como atividade prática durante o curso Técnico em Desenvolvimento de Sistemas da ETEC.

A proposta foi explorar possibilidades de automação utilizando Python, trabalhando com leitura de arquivos Excel, manipulação de dados, integração com serviços externos e controle automatizado de interfaces gráficas.

O projeto possui caráter experimental e representa uma etapa de aprendizado sobre automação de processos.

## Observações

- O WhatsApp Web precisa estar autenticado no navegador.
- A planilha precisa estar no diretório correto.
- A automação depende da disponibilidade e do comportamento da interface do navegador.
- O projeto não possui tratamento completo de erros ou validação de números.

**Importante:** utilize automações de mensagens somente com destinatários que tenham consentido em recebê-las, respeitando as regras do WhatsApp e a legislação aplicável. Este projeto não deve ser utilizado para envio de mensagens em massa não solicitadas.

---

**Projeto desenvolvido para fins educacionais.**
