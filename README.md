
# Conversor Excel para CSV

Este projeto fornece uma interface gráfica simples que permite converter arquivos Excel (`.xlsx` ou `.xls`) para o formato CSV, com valores numéricos formatados corretamente e colunas específicas (como a alíquota) convertidas para percentuais completos.

## Funcionalidades

- Remove espaços em branco nos valores numéricos.
- Converte valores numéricos do estilo "1.000,00" para "1000.00".
- Multiplica a coluna de alíquota por 100 caso seu valor seja menor que 1, convertendo para percentual completo.
- Salva o arquivo convertido em formato CSV com delimitador `;` e sem cabeçalho.

## Requisitos

- Python 3.7 ou superior
- `pandas`
- `openpyxl` (para trabalhar com arquivos Excel)
- `tkinter` (geralmente incluído por padrão no Python)

## Como rodar o projeto localmente

1. Clone o repositório:

```bash
git clone https://github.com/usuario/excel-csv-converter.git
cd excel-csv-converter
```

2. Crie um ambiente virtual (opcional, mas recomendado):

```bash
python -m venv venv
source venv/bin/activate  # Para Linux ou MacOS
venv\Scripts\activate  # Para Windows
```

3. Instale as dependências:

```bash
pip install -r requirements.txt
```

> Certifique-se de que o arquivo `requirements.txt` contenha o seguinte:
> ```
> pandas
> openpyxl
> ```

4. Execute o programa:

```bash
python main.py
```

A interface gráfica será aberta, permitindo que você selecione um arquivo Excel e um local para salvar o arquivo CSV gerado.

## Como gerar um executável

Para criar um executável do programa que pode ser distribuído para outros usuários (que não têm Python instalado), siga os passos abaixo:

1. Instale o `pyinstaller`:

```bash
pip install pyinstaller
```

2. Gere o executável com o seguinte comando:

```bash
pyinstaller --onefile --windowed main.py
```

Explicação dos argumentos:
- `--onefile`: Empacota todo o programa em um único arquivo executável.
- `--windowed`: Remove a exibição do terminal/console (opcional, para interfaces gráficas).

3. Após a execução do comando, o executável será gerado na pasta `dist/` do projeto. O arquivo estará pronto para ser distribuído e executado em outros computadores.

## Estrutura do projeto

```
excel-csv-converter/
│
├── main.py              # Arquivo principal com a lógica do programa
├── requirements.txt     # Arquivo com dependências do projeto
├── README.md            # Arquivo de instruções do projeto
├── dist/                # Pasta onde o executável será salvo
├── build/               # Gerado pelo PyInstaller (pode ser ignorado)
└── __pycache__/         # Cache do Python (pode ser ignorado)
```

## Problemas conhecidos

- Certifique-se de que os arquivos Excel estão formatados corretamente, e que a coluna de alíquota contém valores decimais entre 0 e 1 se quiser que eles sejam convertidos para percentuais.

## Contribuição

Fique à vontade para abrir _issues_ ou enviar _pull requests_ para melhorias ou correções.