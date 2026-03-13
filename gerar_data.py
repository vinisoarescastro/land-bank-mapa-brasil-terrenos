#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║           GERADOR DE data.js — Land Bank Grupo Brasil            ║
║  Lê a planilha Excel + arquivos KML e gera o data.js do site.    ║
╚══════════════════════════════════════════════════════════════════╝

USO:
    python gerar_data.py

DEPENDÊNCIAS:
    pip install openpyxl lxml

CONFIGURAÇÃO (edite a seção abaixo):
"""

# ─────────────────────────────────────────────
#   ⚙️  CONFIGURAÇÃO — EDITE AQUI
# ─────────────────────────────────────────────

# Caminho da planilha Excel
EXCEL_PATH = "areas_land_bank_com_id.xlsx"

# Nome da aba/sheet da planilha (None = primeira aba)
EXCEL_SHEET = None

# Pasta onde estão os arquivos .kml (busca recursiva em subpastas)
KML_FOLDER = "kml"

# Arquivo de saída
OUTPUT_PATH = "data.js"

# ── Vinculação por ID ──────────────────────────────────────────
# Coluna do Excel que contém o ID único do empreendimento (ex: MAP113).
# Esse valor deve coincidir com o prefixo do nome do arquivo KML.
COLUNA_ID = "ID"

# Expressão regular para extrair o ID do nome do arquivo KML.
# Padrão "MAP" seguido de dígitos: MAP106, MAP113, MAP200...
# Ajuste se o seu padrão de IDs for diferente (ex: r'^(AREA\d+)').
ID_REGEX = r'^(MAP\d+)'

# Mapeamento das colunas da planilha para os campos do sistema.
# Os nomes à direita devem ser EXATAMENTE iguais aos cabeçalhos da planilha.
COLUNAS = {
    "nome":                "Nome",
    "codigo":              "Código",
    "regional":            "Regional",
    "cidade":              "Cidade",
    "empreendimento":      "Empreendimento",
    "tipo":                "Tipo",
    "year":                "Year",
    "on_off":              "[ON / OFF]",
    "area_total":          "Area Total",
    "total_unidades":      "Total de Unidades",
    "vgv_total":           "VGV Total\n(R$mm)",
    "vgv_bt":              "VGV Total\n(R$mm) BT",
    "custo_terreno":       "Custo Total do Terreno\n(Pré Rateio - R$mm)",
    "custo_construcao":    "Custo de Construção\n(Pré Rateio - R$mm)",
    "participacao_buriti": "Participação Buriti",
    "data_lancamento":     "Data de Lançamento",
}

# Cores por regional (adicione/edite conforme necessário)
CORES = {
    "NORTE":           "#c0392b",
    "NORDESTE I":      "#d35400",
    "NORDESTE II":     "#e67e22",
    "CENTRO OESTE":    "#27ae60",
    "CENTRO OESTE II": "#16a085",
    "SUDESTE":         "#2980b9",
    "SUL":             "#8e44ad",
    "TOCANTINS":       "#0e6655",
    "OESTE":           "#c2185b",
    "None":            "#7f8c8d",
}

# ─────────────────────────────────────────────
#   🔧  CÓDIGO — NÃO É NECESSÁRIO EDITAR
# ─────────────────────────────────────────────

import os
import re
import json
import math
import unicodedata
from pathlib import Path
from datetime import datetime, date


def normalizar(texto):
    """Remove acentos, espaços extras e deixa em maiúsculas para comparação."""
    if not texto:
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r'\.KML$', '', texto)
    texto = Path(texto).name
    return texto


def extrair_id(texto, regex=ID_REGEX):
    """
    Extrai o ID do início de um texto usando o regex configurado.

    Exemplos:
        "MAP113 - VARZEA GRANDE - GARDEM PRIME I E II" → "MAP113"
        "MAP106 - RIO BRANCO - CIDADE JARDIM II"       → "MAP106"
        "CHÁCARA PARAISO"                              → None  (sem ID reconhecível)
    """
    if not texto:
        return None
    m = re.match(regex, str(texto).strip(), re.IGNORECASE)
    return m.group(1).upper() if m else None


def ler_excel(path, sheet=None):
    """Lê o Excel e retorna (lista_de_registros, cabecalho)."""
    try:
        import openpyxl
    except ImportError:
        raise ImportError("Execute: pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb.active

    linhas = list(ws.iter_rows(values_only=True))
    if not linhas:
        return [], []

    cabecalho = [str(c).strip() if c is not None else "" for c in linhas[0]]
    registros = []

    for linha in linhas[1:]:
        if all(v is None for v in linha):
            continue
        registro = {}
        for i, col in enumerate(cabecalho):
            registro[col] = linha[i] if i < len(linha) else None
        registros.append(registro)

    print(f"  ✅ Excel: {len(registros)} linhas lidas — colunas: {cabecalho}")
    return registros, cabecalho


def extrair_coordenadas_kml(caminho_kml):
    """
    Extrai de um arquivo KML:
      - nome_kml : valor da tag <name> do Document (ou Placemark raiz)
      - poligonos: lista de polígonos [[lat, lng], ...]
      - centroide: [lat_media, lng_media]

    Retorna: (nome_kml, poligonos, centroide)
    """
    try:
        from lxml import etree
    except ImportError:
        raise ImportError("Execute: pip install lxml")

    try:
        tree = etree.parse(caminho_kml)
    except Exception as e:
        print(f"  ⚠️  Erro ao parsear {caminho_kml}: {e}")
        return None, [], None

    root = tree.getroot()

    # Remove namespace para facilitar busca
    for elem in root.iter():
        if elem.tag.startswith("{"):
            elem.tag = elem.tag.split("}", 1)[1]

    # ── Extrai <name> ────────────────────────────────────────────
    nome_kml = None
    doc = root.find(".//Document")
    if doc is not None:
        tag = doc.find("name")
        if tag is not None and tag.text and tag.text.strip():
            nome_kml = tag.text.strip()

    if not nome_kml:
        pm = root.find(".//Placemark")
        if pm is not None:
            tag = pm.find("name")
            if tag is not None and tag.text and tag.text.strip():
                nome_kml = tag.text.strip()

    if not nome_kml:
        tag = root.find(".//name")
        if tag is not None and tag.text and tag.text.strip():
            nome_kml = tag.text.strip()

    # ── Extrai polígonos ──────────────────────────────────────────
    poligonos = []
    todos_lats = []
    todos_lngs = []

    for polygon in root.iter("Polygon"):
        for coords_tag in polygon.iter("coordinates"):
            texto = coords_tag.text
            if not texto:
                continue
            pontos = []
            for token in texto.strip().split():
                partes = token.split(",")
                if len(partes) >= 2:
                    try:
                        lng = float(partes[0])
                        lat = float(partes[1])
                        pontos.append([lat, lng])
                        todos_lats.append(lat)
                        todos_lngs.append(lng)
                    except ValueError:
                        pass
            if pontos:
                poligonos.append(pontos)

    # Fallback: LineString
    if not poligonos:
        for ls in root.iter("LineString"):
            for coords_tag in ls.iter("coordinates"):
                texto = coords_tag.text
                if not texto:
                    continue
                pontos = []
                for token in texto.strip().split():
                    partes = token.split(",")
                    if len(partes) >= 2:
                        try:
                            lng = float(partes[0])
                            lat = float(partes[1])
                            pontos.append([lat, lng])
                            todos_lats.append(lat)
                            todos_lngs.append(lng)
                        except ValueError:
                            pass
                if pontos:
                    poligonos.append(pontos)

    centroide = None
    if todos_lats and todos_lngs:
        centroide = [sum(todos_lats) / len(todos_lats), sum(todos_lngs) / len(todos_lngs)]

    return nome_kml, poligonos, centroide


def calcular_area(poligonos):
    """Estima área total em m² usando fórmula de Shoelace (aproximação plana)."""
    total = 0.0
    for pol in poligonos:
        if len(pol) < 3:
            continue
        lat_ref = pol[0][0]
        escala_lat = 111320.0
        escala_lng = 111320.0 * math.cos(math.radians(lat_ref))
        coords_m = [(p[1] * escala_lng, p[0] * escala_lat) for p in pol]
        n = len(coords_m)
        area = 0.0
        for i in range(n):
            j = (i + 1) % n
            area += coords_m[i][0] * coords_m[j][1]
            area -= coords_m[j][0] * coords_m[i][1]
        total += abs(area) / 2.0
    return total


def serializar_valor(v):
    """Converte tipos Python para tipos serializáveis em JSON."""
    if isinstance(v, (datetime, date)):
        return str(v)
    if isinstance(v, float) and math.isnan(v):
        return None
    return v


def construir_e(reg):
    """Monta o dicionário 'e' de um item a partir de um registro do Excel."""
    e = {}
    for campo_sistema, col_excel in COLUNAS.items():
        if col_excel is None:
            e[campo_sistema] = None
            continue
        valor = None
        for k, v in reg.items():
            if k.strip().lower() == col_excel.strip().lower():
                valor = serializar_valor(v)
                break
        e[campo_sistema] = valor
    return e


def main():
    print("\n" + "═" * 60)
    print("  🗺️  GERADOR data.js — Land Bank Grupo Brasil")
    print("═" * 60)

    # ── 1. Ler Excel ──────────────────────────────────────────────
    print(f"\n📊 Lendo planilha: {EXCEL_PATH}")
    if not os.path.exists(EXCEL_PATH):
        print(f"  ❌ Arquivo não encontrado: {EXCEL_PATH}")
        print(f"     Verifique o caminho em EXCEL_PATH no início do script.")
        return

    registros_excel, cabecalho_excel = ler_excel(EXCEL_PATH, EXCEL_SHEET)

    # ── 2. Montar índice Excel: ID → [lista de registros] ─────────
    # Suporta múltiplos registros com o mesmo ID (IDs repetidos na planilha).
    # Cada registro da lista será vinculado ao mesmo arquivo KML.
    indice_excel = {}   # { "MAP113": [reg_a, reg_b, ...], "MAP106": [reg_c], ... }
    col_id_real = None

    for col in cabecalho_excel:
        if col.strip().lower() == COLUNA_ID.strip().lower():
            col_id_real = col
            break

    if not col_id_real:
        print(f"\n  ⚠️  Coluna de ID '{COLUNA_ID}' não encontrada no Excel.")
        print(f"     Colunas disponíveis: {cabecalho_excel}")
        print(f"     Ajuste COLUNA_ID no início do script.")
        print(f"     Continuando sem vincular dados da planilha...\n")
    else:
        for reg in registros_excel:
            id_val = str(reg.get(col_id_real, "") or "").strip().upper()
            if id_val:
                indice_excel.setdefault(id_val, []).append(reg)

        total_ids      = len(indice_excel)
        ids_duplicados = {k: v for k, v in indice_excel.items() if len(v) > 1}
        print(f"  ✅ {total_ids} IDs únicos no índice")

        if ids_duplicados:
            print(f"  ℹ️  {len(ids_duplicados)} ID(s) com múltiplos registros (correto — um KML, várias linhas):")
            for id_k, regs in sorted(ids_duplicados.items()):
                print(f"     • {id_k}: {len(regs)} registros")

    # ── 3. Ler KMLs ───────────────────────────────────────────────
    print(f"\n📁 Buscando KMLs em: {KML_FOLDER}")
    if not os.path.exists(KML_FOLDER):
        print(f"  ❌ Pasta não encontrada: {KML_FOLDER}")
        print(f"     Verifique o caminho em KML_FOLDER no início do script.")
        return

    arquivos_kml = list(Path(KML_FOLDER).rglob("*.kml")) + list(Path(KML_FOLDER).rglob("*.KML"))
    arquivos_kml = sorted(set(arquivos_kml))
    print(f"  ✅ {len(arquivos_kml)} arquivos KML encontrados")

    # ── 4. Vincular KMLs com registros Excel ──────────────────────
    print(f"\n🔗 Vinculando KMLs com a planilha...")
    print(f"   Chave: prefixo '{ID_REGEX}' do nome do arquivo KML → coluna '{COLUNA_ID}' do Excel\n")

    items = []
    kml_vinculados       = 0   # KMLs que encontraram ao menos 1 registro
    kml_sem_vinculo      = 0   # KMLs sem nenhum registro na planilha
    kml_sem_poligono     = 0   # KMLs sem geometria
    kml_sem_id           = 0   # KMLs cujo nome não contém ID reconhecível
    total_itens_criados  = 0   # Itens criados a partir de KMLs (1 por registro vinculado)
    nao_vinculados       = []  # Lista para o relatório final
    ids_kml_processados  = set()

    for kml_path in arquivos_kml:
        nome_arquivo = kml_path.stem   # ex: "MAP113 - VARZEA GRANDE - GARDEM PRIME I E II"
        nome_kml_tag, poligonos, centroide = extrair_coordenadas_kml(str(kml_path))

        if not poligonos:
            kml_sem_poligono += 1

        # Nome para exibição: prefere a tag <name> interna, cai no nome do arquivo
        nome_display = nome_kml_tag or nome_arquivo

        # Extrai o ID: tenta primeiro no nome do arquivo, depois na tag <name>
        id_kml = extrair_id(nome_arquivo)
        if not id_kml and nome_kml_tag:
            id_kml = extrair_id(nome_kml_tag)

        if not id_kml:
            kml_sem_id += 1

        # Busca registros vinculados no índice Excel
        registros_vinculados = indice_excel.get(id_kml, []) if id_kml else []

        if registros_vinculados:
            # ✅ Vinculado: cria um item para CADA registro com esse ID.
            # Todos compartilham o mesmo polígono e centroide do KML único.
            kml_vinculados += 1
            ids_kml_processados.add(id_kml)
            for reg in registros_vinculados:
                items.append({
                    "id": id_kml,
                    "n":  nome_display,
                    "p":  poligonos,
                    "c":  centroide,
                    "e":  construir_e(reg),
                })
                total_itens_criados += 1
        else:
            # ❌ Sem vínculo: cria um item sem dados da planilha
            kml_sem_vinculo += 1
            nao_vinculados.append((nome_display, id_kml or "sem ID", str(kml_path)))
            items.append({
                "id": id_kml,
                "n":  nome_display,
                "p":  poligonos,
                "c":  centroide,
                "e":  None,
            })
            total_itens_criados += 1

    # ── 5. Registros Excel sem KML ────────────────────────────────
    # IDs que existem na planilha mas não têm arquivo KML correspondente.
    sem_kml = 0
    if col_id_real:
        for id_val, regs in indice_excel.items():
            if id_val not in ids_kml_processados:
                for reg in regs:
                    items.append({
                        "id": id_val,
                        "n":  reg.get(col_id_real, id_val),
                        "p":  [],
                        "c":  None,
                        "e":  construir_e(reg),
                    })
                    sem_kml += 1

    # ── 6. Identificar itens sem localização ──────────────────────
    # Um item é considerado "sem localização" quando não possui polígono
    # nem centroide — ou seja, não há nenhuma coordenada associada a ele.
    # Isso ocorre em dois casos:
    #   a) KML sem geometria válida (arquivo vazio ou corrompido)
    #   b) Registro da planilha sem arquivo KML correspondente
    sem_localizacao = []
    for item in items:
        if not item["p"] and not item["c"]:
            e = item["e"] or {}
            sem_localizacao.append({
                "id":       item.get("id") or "—",
                "nome":     e.get("nome") or item["n"] or "—",
                "cidade":   e.get("cidade") or "—",
                "regional": e.get("regional") or "—",
                "motivo":   "sem KML" if not item["p"] and item["e"] else "KML sem geometria",
            })

    # ── 7. Calcular estatísticas ───────────────────────────────────
    # VGV, unidades e área são somados de registros_excel diretamente
    # para incluir TODOS os registros, inclusive os de IDs duplicados.
    on_map     = sum(1 for i in items if i["p"])
    total_area = sum(calcular_area(i["p"]) for i in items if i["p"])

    col_vgv    = COLUNAS.get("vgv_total")
    col_vgv_bt = COLUNAS.get("vgv_bt")
    col_units  = COLUNAS.get("total_unidades")

    def _soma_excel(col):
        total = 0.0
        for reg in registros_excel:
            for k, v in reg.items():
                if k.strip().lower() == col.strip().lower():
                    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
                        total += v
                    break
        return total

    total_vgv      = _soma_excel(col_vgv)    if col_vgv    else 0.0
    total_vgv_bt   = _soma_excel(col_vgv_bt) if col_vgv_bt else 0.0
    total_unidades = _soma_excel(col_units)  if col_units  else 0.0

    stats = {
        "total":        len(items),
        "on_map":       on_map,
        "total_units":  total_unidades,
        "total_area":   round(total_area, 0),
        "total_vgv":    round(total_vgv, 2),
        "total_vgv_bt": round(total_vgv_bt, 2),
    }

    # Resumo por regional
    regional_summary = {}
    for item in items:
        if item["e"] and item["e"].get("regional"):
            r = item["e"]["regional"]
            if r not in regional_summary:
                regional_summary[r] = {"count": 0, "units": 0, "vgv": 0}
            regional_summary[r]["count"] += 1
            regional_summary[r]["units"] += item["e"].get("total_unidades") or 0
            regional_summary[r]["vgv"]   += item["e"].get("vgv_total") or 0

    # ── 8. Escrever data.js ────────────────────────────────────────
    data = {
        "items":            items,
        "colors":           CORES,
        "stats":            stats,
        "regional_summary": regional_summary,
    }

    js_content = "const DATA = " + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ";"

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(js_content)

    # ── 9. Relatório final ─────────────────────────────────────────
    print(f"{'═' * 60}")
    print(f"  ✅ data.js gerado com sucesso!")
    print(f"{'═' * 60}")
    print(f"  📦 Total de itens gerados:            {len(items)}")
    print(f"  🗺️  Com polígono KML:                  {on_map}")
    print(f"  📍 Sem localização:                   {len(sem_localizacao)}")
    print(f"  🔗 KMLs vinculados à planilha:        {kml_vinculados}")
    print(f"  ❌ KMLs sem vínculo na planilha:      {kml_sem_vinculo}")
    print(f"  📋 Registros só na planilha (sem KML):{sem_kml}")
    print(f"  ⚠️  KMLs sem geometria:               {kml_sem_poligono}")
    print(f"  🏷️  KMLs sem ID reconhecível:         {kml_sem_id}")
    print(f"  💰 VGV Total:                         R$ {total_vgv:,.1f} mi")
    print(f"  🏘️  Total de unidades:                 {total_unidades:,.0f}")
    print(f"\n  📄 Arquivo gerado: {os.path.abspath(OUTPUT_PATH)}")

    if sem_localizacao:
        print(f"\n  📍 {len(sem_localizacao)} registro(s) SEM localização (sem polígono nem centroide):")
        col_id_w     = max(len(r["id"])       for r in sem_localizacao)
        col_nome_w   = max(len(r["nome"])     for r in sem_localizacao)
        col_cidade_w = max(len(r["cidade"])   for r in sem_localizacao)
        col_reg_w    = max(len(r["regional"]) for r in sem_localizacao)
        header = (
            f"     {'ID':<{col_id_w}}  "
            f"{'Nome':<{col_nome_w}}  "
            f"{'Cidade':<{col_cidade_w}}  "
            f"{'Regional':<{col_reg_w}}  Motivo"
        )
        print(header)
        print("     " + "─" * (len(header) - 5))
        for r in sorted(sem_localizacao, key=lambda x: (x["regional"], x["id"])):
            print(
                f"     {r['id']:<{col_id_w}}  "
                f"{r['nome']:<{col_nome_w}}  "
                f"{r['cidade']:<{col_cidade_w}}  "
                f"{r['regional']:<{col_reg_w}}  {r['motivo']}"
            )

    if nao_vinculados:
        print(f"\n  ⚠️  {len(nao_vinculados)} KML(s) SEM vínculo na planilha:")
        for nome, id_encontrado, path in nao_vinculados:
            label = f"ID={id_encontrado}" if id_encontrado != "sem ID" else "sem ID reconhecível"
            print(f"     • [{label}] \"{nome}\"")
            print(f"       → {path}")

    print(f"\n{'═' * 60}")
    print("  👉 Próximo passo: faça commit e push do data.js para")
    print("     o repositório. O site atualizará automaticamente.")
    print(f"{'═' * 60}\n")


if __name__ == "__main__":
    main()