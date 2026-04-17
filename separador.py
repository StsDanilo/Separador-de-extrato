import json
import sys
import unicodedata
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

CATEGORIAS_FILE = "categorias.json"


def carregar_categorias(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def sem_acento(texto):
    """Normaliza para ASCII, ignorando acentos e caracteres corrompidos.
    Isso permite que termos do JSON (UTF-8 correto) batam com valores
    do xlsx que podem ter acentos corrompidos como caracteres de substituição."""
    return unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii").lower()


def eh_entrada(tipo):
    return str(tipo).strip().lower() == "entrada"


def eh_saida(tipo):
    t = str(tipo).strip().lower()
    # O xlsx pode ter o 'í' corrompido, então não dá pra comparar direto "saída"
    return t.startswith("sa") and t.endswith("da")


def linha_corresponde(row_data, condicoes):
    lancamento = row_data.get("lancamento", "")
    detalhes   = row_data.get("detalhes", "")

    for cond in condicoes:
        campo  = cond["campo"]
        termos = cond["termos"]
        modo   = cond.get("modo", "contem")

        if campo == "lancamento":
            texto = lancamento
        elif campo == "detalhes":
            texto = detalhes
        else:
            texto = lancamento + " " + detalhes

        # Compara sem acentos para tolerar corrupção de encoding no xlsx
        texto_norm = sem_acento(texto)
        if modo == "contem":
            if not any(sem_acento(t) in texto_norm for t in termos):
                return False
        elif modo == "exato":
            if not any(sem_acento(t) == texto_norm for t in termos):
                return False

    return True


def categorizar_linha(row_data, categorias):
    for nome, cat in categorias.items():
        if linha_corresponde(row_data, cat["condicoes"]):
            return nome
    return "Outros"


def ler_extrato(path):
    wb = load_workbook(path)
    ws = wb.active

    headers_raw = [str(c.value).strip() if c.value else "" for c in ws[1]]

    col_map = {}
    for i, h in enumerate(headers_raw):
        hl = h.lower()
        if hl.startswith("data"):
            col_map["data"] = i
        elif "tipo" in hl:
            # deve vir antes de "lan" pois "Tipo Lançamento" também contém "lan"
            col_map["tipo"] = i
        elif hl.startswith("lan"):
            col_map["lancamento"] = i
        elif "detalhe" in hl:
            col_map["detalhes"] = i
        elif "doc" in hl:
            col_map["documento"] = i
        elif "valor" in hl:
            col_map["valor"] = i

    # Renomeia cabeçalho da coluna tipo para "Tipo" e abrevia
    headers_saida = list(headers_raw)
    idx_tipo_header = col_map.get("tipo")
    if idx_tipo_header is not None:
        headers_saida[idx_tipo_header] = "Tipo"

    entradas = []
    saidas   = []
    saldo_anterior = None

    idx_lancamento = col_map.get("lancamento", 1)
    idx_tipo       = col_map.get("tipo", 5)

    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row)
        row = [str(v).strip() if v is not None else "" for v in row]

        lancamento = row[idx_lancamento]
        tipo       = row[idx_tipo]

        if "saldo anterior" in lancamento.lower():
            row[idx_tipo] = ""
            saldo_anterior = row
            continue
        if "saldo" in lancamento.lower():
            continue

        if eh_entrada(tipo):
            row[idx_tipo] = "(+)"
            entradas.append(row)
        elif eh_saida(tipo):
            row[idx_tipo] = "(-)"
            saidas.append(row)

    return headers_saida, col_map, saldo_anterior, entradas, saidas


def agrupar_por_categoria(linhas, col_map, categorias):
    grupos = {nome: [] for nome in categorias}
    grupos["Outros"] = []

    idx_lanc = col_map.get("lancamento", 1)
    idx_det  = col_map.get("detalhes", 2)

    for row in linhas:
        row_data = {
            "lancamento": row[idx_lanc],
            "detalhes":   row[idx_det],
        }
        cat = categorizar_linha(row_data, categorias)
        grupos[cat].append(row)

    return grupos


FILL_DIA_A  = PatternFill("solid", fgColor="FFFFFF")  # branco
FILL_DIA_B  = PatternFill("solid", fgColor="EFEFEF")  # cinza muito claro
FILL_BRANCO = PatternFill("solid", fgColor="FFFFFF")


def borda_fina():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)


def _aplicar_merged_row(ws, row_idx, num_cols, value, font, alignment, fill):
    """Aplica estilo a todas as células de uma linha mesclada.
    Células não-master em merged ranges precisam de fill explícito,
    senão o Excel exibe o fundo padrão cinza nelas."""
    ws.merge_cells(start_row=row_idx, start_column=1,
                   end_row=row_idx, end_column=num_cols)
    borda = borda_fina()
    borda_interna = Border(top=borda.top, bottom=borda.bottom)
    for c in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill   = fill
        cell.border = borda_interna
    # Bordas externas apenas nas células das pontas
    ws.cell(row=row_idx, column=1).border = Border(
        left=borda.left, right=Side(style=None),
        top=borda.top, bottom=borda.bottom)
    ws.cell(row=row_idx, column=num_cols).border = Border(
        left=Side(style=None), right=borda.right,
        top=borda.top, bottom=borda.bottom)
    master = ws.cell(row=row_idx, column=1)
    master.value     = value
    master.font      = font
    master.alignment = alignment


def escrever_header_bloco(ws, row_idx, titulo, num_cols):
    _aplicar_merged_row(
        ws, row_idx, num_cols, titulo,
        Font(bold=True, size=13),
        Alignment(horizontal="center", vertical="center"),
        FILL_BRANCO,
    )
    ws.row_dimensions[row_idx].height = 22


def escrever_header_categoria(ws, row_idx, nome, num_cols):
    _aplicar_merged_row(
        ws, row_idx, num_cols, f"  {nome}",
        Font(bold=True, size=10),
        Alignment(horizontal="left", vertical="center"),
        FILL_BRANCO,
    )
    ws.row_dimensions[row_idx].height = 18


def escrever_header_colunas(ws, row_idx, headers):
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=row_idx, column=c_idx, value=h)
        cell.font      = Font(bold=True, size=9)
        cell.alignment = Alignment(horizontal="center")
        cell.border    = borda_fina()
        cell.fill      = FILL_BRANCO


def escrever_linha_dados(ws, row_idx, row, col_map=None, fill=None):
    borda    = borda_fina()
    idx_val  = col_map.get("valor")      if col_map else None
    idx_lanc = col_map.get("lancamento") if col_map else None
    idx_doc  = col_map.get("documento")  if col_map else None
    for c_idx, val in enumerate(row, start=1):
        col0 = c_idx - 1
        cell = ws.cell(row=row_idx, column=c_idx, value=val)
        cell.font   = Font(size=9)
        cell.border = borda
        if fill:
            cell.fill = fill
        if col0 == idx_val:
            cell.alignment = Alignment(horizontal="right")
        elif col0 in (idx_lanc, idx_doc):
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def linha_em_branco(ws, row_idx, num_cols):
    for c in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.value  = None
        cell.fill   = PatternFill(fill_type=None)
        cell.border = Border()


def escrever_bloco(ws, current_row, titulo, grupos, headers, col_map, num_cols):
    escrever_header_bloco(ws, current_row, titulo, num_cols)
    current_row += 1

    idx_data  = col_map.get("data", 0)
    idx_valor = col_map.get("valor")

    primeiro = True
    for nome_cat, linhas in grupos.items():
        if not linhas:
            continue

        if not primeiro:
            linha_em_branco(ws, current_row, num_cols)
            current_row += 1

        escrever_header_categoria(ws, current_row, nome_cat, num_cols)
        current_row += 1

        escrever_header_colunas(ws, current_row, headers)
        current_row += 1

        data_atual = None
        fill_par   = False
        for row in linhas:
            data_linha = row[idx_data] if idx_data < len(row) else ""
            if data_linha != data_atual:
                data_atual = data_linha
                fill_par   = not fill_par
            fill = FILL_DIA_A if fill_par else FILL_DIA_B
            escrever_linha_dados(ws, current_row, row, col_map=col_map, fill=fill)
            current_row += 1

        primeiro = False

    return current_row


def gerar_extrato(input_path, categorias_path=CATEGORIAS_FILE):
    categorias = carregar_categorias(categorias_path)
    headers, col_map, saldo_anterior, entradas, saidas = ler_extrato(input_path)

    num_cols = len(headers)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_path = f"Extrato_{timestamp}.xlsx"

    wb_out = Workbook()
    ws = wb_out.active
    ws.title = "Extrato Categorizado"

    current_row = 1

    if saldo_anterior:
        escrever_linha_dados(ws, current_row, saldo_anterior, col_map=col_map)
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=9)
        current_row += 1

    linha_em_branco(ws, current_row, num_cols)
    current_row += 1

    grupos_entrada = agrupar_por_categoria(entradas, col_map, categorias)
    current_row = escrever_bloco(ws, current_row, "ENTRADAS", grupos_entrada, headers, col_map, num_cols)

    linha_em_branco(ws, current_row, num_cols)
    current_row += 1
    linha_em_branco(ws, current_row, num_cols)
    current_row += 1

    grupos_saida = agrupar_por_categoria(saidas, col_map, categorias)
    current_row = escrever_bloco(ws, current_row, "SAÍDAS", grupos_saida, headers, col_map, num_cols)

    # Larguras fixas para colunas com conteúdo previsível
    larguras_fixas = {
        col_map.get("data"):       11,
        col_map.get("lancamento"): 18,
        col_map.get("valor"):      10,
        col_map.get("documento"):  15,
    }
    for col in ws.columns:
        col_idx    = col[0].column - 1
        col_letter = get_column_letter(col[0].column)
        if col_idx in larguras_fixas:
            ws.column_dimensions[col_letter].width = larguras_fixas[col_idx]
            continue
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

    wb_out.save(output_path)
    print(f"Arquivo gerado: {output_path}")
    return output_path


if __name__ == "__main__":
    entrada = sys.argv[1] if len(sys.argv) > 1 else "ExtratoTeste.xlsx"
    cats    = sys.argv[2] if len(sys.argv) > 2 else CATEGORIAS_FILE
    gerar_extrato(entrada, cats)
