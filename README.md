# Comparador de Concluintes BFD

Script Python que compara a lista de alunos concluintes de uma planilha Excel com as pastas de evidências baixadas do Google Drive, identificando divergências.

## Pré-requisitos

- Python 3.8 ou superior
- Biblioteca `openpyxl`

```bash
pip install openpyxl
```

## Configuração

Abra o arquivo `comparar_concluintes.py` e edite as constantes no início do arquivo:

```python
EXCEL_PATH = r"C:\caminho\para\sua\planilha.xlsx"
SHEET_NAME = "NomeDaAba"
ESTADOS    = ["Pernambuco"]
DRIVE_ROOT = r"C:\caminho\para\pasta\do\drive"
```

| Constante | Descrição |
|---|---|
| `EXCEL_PATH` | Caminho completo para o arquivo `.xlsx` |
| `SHEET_NAME` | Nome da aba da planilha a ser lida |
| `ESTADOS` | Lista de estados a filtrar (deve bater com a coluna "Estado" da planilha) |
| `DRIVE_ROOT` | Pasta raiz com as subpastas de turmas e alunos |

### Estrutura esperada da planilha

A aba deve ter colunas na seguinte ordem (sem cabeçalho obrigatório):

| Coluna A | Coluna B | Coluna C |
|---|---|---|
| Estado | Turma | Nome do aluno |

### Estrutura esperada das pastas no Drive

```
DRIVE_ROOT/
  T01PEC1/
    joao_silva/
    maria_souza/
  T02PEC1/
    ...
```

> O prefixo `Turma` nas pastas é normalizado automaticamente (ex: `Turma01PEC1` → `T01PEC1`).

## Execução

```bash
python comparar_concluintes.py
```

## Exemplo de saída

```
=================================================================
❌  PASTAS NO DRIVE QUE NÃO ESTÃO NA PLANILHA
    (candidatas a remoção)
=================================================================

  T01PEC1:
    - jose_ferreira_teste

  Total: 1 pasta(s) extra(s)

=================================================================
⚠️   ALUNOS NA PLANILHA SEM PASTA NO DRIVE
    (evidências ausentes)
=================================================================

  T01PEC1:
    - Ana Beatriz Rodrigues Lima

  Total: 1 aluno(s) sem pasta
```
