# Separador de Extrato Bancário

Projeto simples de uso próprio para resolver um problema do dia a dia: organizar o extrato bancário do **Banco do Brasil** (`.xlsx`) separando os lançamentos por categorias personalizadas e dividindo entradas e saídas em blocos distintos.

## O problema

O extrato exportado pelo Banco do Brasil lista todos os lançamentos em ordem cronológica, misturando entradas e saídas. Para facilitar a análise mensal, precisava de uma visualização organizada por tipo de lançamento, com agrupamentos que fizessem sentido para o meu fluxo financeiro.

## O que o script faz

1. Lê o arquivo `.xlsx` do extrato
2. Separa todos os lançamentos em **ENTRADAS** e **SAÍDAS**
3. Dentro de cada bloco, agrupa os lançamentos por categorias definidas no `categorias.json`
4. Lançamentos que não se encaixam em nenhuma categoria vão para um grupo **Outros**
5. Gera um novo arquivo `.xlsx` com nome automático baseado na data/hora (`Extrato_AAAA-MM-DD_HH-MM.xlsx`), pronto para impressão monocromática

## Estrutura do arquivo de saída

```
Saldo Anterior

ENTRADAS
  ├── [Categoria A]
  │     Data | Lançamento | Detalhes | Nº doc | Valor | Tipo
  │     ...
  ├── [Categoria B]
  │     ...
  └── Outros
        ...

SAÍDAS
  ├── [Categoria C]
  │     ...
  └── Outros
        ...
```

- Dias alternados têm fundo levemente diferente para facilitar a leitura
- Categorias sem lançamentos no período simplesmente não aparecem

## Como usar

```bash
python separador.py <arquivo_extrato.xlsx> [categorias.json]
```

**Exemplos:**
```bash
# Usando o categorias.json padrão
python separador.py extrato_marco.xlsx

# Especificando outro arquivo de categorias
python separador.py extrato_marco.xlsx minhas_categorias.json
```

## Dependências

```bash
pip install openpyxl
```

## Configurando as categorias

Edite o arquivo `categorias.json` para definir seus próprios agrupamentos. Cada categoria tem uma lista de condições que são verificadas contra as colunas **Lançamento** e **Detalhes** do extrato.

```json
{
  "Nome da Categoria": {
    "condicoes": [
      { "campo": "lancamento", "termos": ["Texto A", "Texto B"], "modo": "contem" }
    ]
  }
}
```

**Campos disponíveis:**
- `campo`: `"lancamento"`, `"detalhes"` ou qualquer outro valor para buscar nos dois
- `termos`: lista de strings — basta uma delas estar presente para a condição ser satisfeita
- `modo`: `"contem"` (padrão) ou `"exato"`

A busca é insensível a maiúsculas/minúsculas e ignora acentos, o que garante compatibilidade mesmo quando o arquivo `.xlsx` apresenta problemas de codificação de caracteres.

Uma categoria pode ter **múltiplas condições** — todas precisam ser satisfeitas (AND lógico).

**Exemplo com múltiplas condições:**
```json
{
  "PIX para Fornecedor X": {
    "condicoes": [
      { "campo": "lancamento", "termos": ["Pix - Enviado"], "modo": "contem" },
      { "campo": "detalhes",   "termos": ["FORNECEDOR X"],  "modo": "contem" }
    ]
  }
}
```

## Arquivos de exemplo

| Arquivo | Descrição |
|---|---|
| `ExtratoExemplo.xlsx` | Modelo de extrato de entrada com dados fictícios |
| `ExtratoExemploSaida.xlsx` | Resultado gerado a partir do exemplo |
| `categorias_exemplo.json` | Exemplo de arquivo de categorias |
