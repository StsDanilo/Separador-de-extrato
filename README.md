# Separador de Extratos

Aplicativo unificado para organizar extratos do Banco do Brasil, Mercado Pago e Stone sem usar o console.

## Recursos

- Interface desktop com CustomTkinter.
- Escolha manual do banco.
- Detecção automática do banco ao selecionar o arquivo.
- Processamento dos formatos antigos do Banco do Brasil, Mercado Pago e Stone.
- Pasta de saída configurável.
- Histórico dos últimos arquivos usados.
- Editor visual de categorias do Banco do Brasil.
- Botões para abrir o arquivo gerado ou a pasta de saída.
- Script de build com PyInstaller.

## Como executar em modo desenvolvimento

```powershell
.\run.ps1
```

## Como gerar o executável

Dentro da pasta deste projeto:

```powershell
.\build.ps1
```

O executável será criado em:

```text
dist\SeparadorExtratos.exe
```

O build precisa ser feito com uma instalação normal do Python para Windows com Tcl/Tk disponível. O Python embutido do ambiente Codex pode gerar um `.exe` sem `tkinter`, que não abre a interface.

## Categorias do Banco do Brasil

O app copia a configuração inicial de `config/categorias.json` para a pasta de dados do usuário na primeira execução. Depois disso, as alterações feitas no editor visual são salvas no arquivo do usuário, sem mexer no arquivo original do projeto.

No Windows, esse arquivo fica em:

```text
%APPDATA%\SeparadorExtratos\categorias.json
```

## Regras da Stone

O extrato Stone é dividido em `CRÉDITOS` e `DÉBITOS`.

- `Aplicação na Reserva Stone`: movimentação de débito com destino `Desconhecido`.
- `Resgate da Reserva Stone`: movimentação de crédito com origem `Desconhecido`.
- `Débitos comuns`: movimentação de débito com destino diferente de `Desconhecido`.
- `Créditos comuns`: movimentação de crédito com origem diferente de `Desconhecido`.

A planilha gerada mantém apenas as colunas `Destino`, `Data`, `Valor` e `Saldo depois`, nessa ordem (as colunas `Movimentação`, `Tipo`, `Tarifa`, `Situação`, `Origem`, `Origem Documento`, `Origem Instituição`, `Destino Documento`, `Destino Instituição` e `Saldo antes` não entram na saída, embora `Movimentação`, `Destino` e `Origem` continuem sendo usadas internamente para classificar os lançamentos).

No topo da planilha é exibido um cabeçalho com o saldo antes do primeiro lançamento do mês e o saldo final após o último lançamento. Os lançamentos são sempre reordenados do mais antigo para o mais recente (começando pelo primeiro dia do mês), mesmo que o extrato original venha do Stone na ordem inversa.

Arquivos Stone `.xlsx` e `.xlsm` são lidos com `openpyxl`. Arquivos `.xls` também são aceitos quando forem planilhas Excel antigas reais, usando `xlrd`. Se o banco entregar um `.xls` que na verdade é HTML renomeado, converta para `.xlsx` antes de processar.

## Observação para desenvolvimento

Use `run.ps1` ou `run_app.py` para iniciar o aplicativo. O arquivo `app.py` concentra a interface, mas o launcher evita problemas de inicialização do Tcl/Tk em alguns ambientes Python.

## Validação

Para validar os processadores usando os arquivos existentes das pastas antigas:

```powershell
.\.venv\Scripts\python.exe tools\validate_existing_files.py
```

As saídas de validação são criadas em `validation_output/`.
