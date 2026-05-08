#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║           GERADOR DE data.js — Land Bank Grupo Brasil            ║
║  Lê a planilha Excel + arquivos KML e gera o data.js do site.    ║
╚══════════════════════════════════════════════════════════════════╝
"""

# ─────────────────────────────────────────────
#   ⚙️  CONFIGURAÇÃO — EDITE AQUI
# ─────────────────────────────────────────────

EXCEL_PATH  = "areas_land_bank_com_id.xlsx"
EXCEL_SHEET = None
KML_FOLDER  = "kml"
OUTPUT_PATH = "data.js"
COLUNA_ID   = "ID"
ID_REGEX    = r'^(MAP\d+)'

COLUNAS = {
    "nome":                "Nome",
    "codigo":              "Código",
    "regional":            "Regional",
    "cidade":              "Cidade",
    "uf":                  "UF",
    "empreendimento":      "Empreendimento",
    "tipo":                "Tipo",
    "year":                "Year",
    "on_off":              "[ON / OFF]",
    "area_total":          "Area Total m2",
    "total_unidades":      "Total de Unidades",
    "vgv_total":           "VGV Total\n(R$mm)",
    "vgv_bt":              "VGV Total\n(R$mm) BT",
    "custo_terreno":       "Custo Total do Terreno\n(Pré Rateio - R$mm)",
    "custo_construcao":    "Custo de Construção\n(Pré Rateio - R$mm)",
    "participacao_buriti": "Participação Buriti",
    "data_lancamento":     "Data de Lançamento",
}

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

import os, re, json, math, subprocess, unicodedata
from pathlib import Path
from datetime import datetime, date


def normalizar(texto):
    if not texto:
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r'\.KML$', '', texto)
    texto = Path(texto).name
    return texto


def extrair_id(texto, regex=ID_REGEX):
    if not texto:
        return None
    m = re.match(regex, str(texto).strip(), re.IGNORECASE)
    return m.group(1).upper() if m else None


def ler_excel(path, sheet=None):
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
    for elem in root.iter():
        if elem.tag.startswith("{"):
            elem.tag = elem.tag.split("}", 1)[1]
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
    poligonos  = []
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
        centroide = [sum(todos_lats)/len(todos_lats), sum(todos_lngs)/len(todos_lngs)]
    return nome_kml, poligonos, centroide


def serializar_valor(v):
    if isinstance(v, (datetime, date)):
        return str(v)
    if isinstance(v, float) and math.isnan(v):
        return None
    return v


def construir_e(reg):
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
    print("\n" + "═"*60)
    print("  🗺️  GERADOR data.js — Land Bank Grupo Brasil")
    print("═"*60)

    print(f"\n📊 Lendo planilha: {EXCEL_PATH}")
    if not os.path.exists(EXCEL_PATH):
        print(f"  ❌ Arquivo não encontrado: {EXCEL_PATH}")
        return

    registros_excel, cabecalho_excel = ler_excel(EXCEL_PATH, EXCEL_SHEET)

    indice_excel = {}
    col_id_real  = None

    for col in cabecalho_excel:
        if col.strip().lower() == COLUNA_ID.strip().lower():
            col_id_real = col
            break

    if not col_id_real:
        print(f"\n  ⚠️  Coluna de ID '{COLUNA_ID}' não encontrada no Excel.")
        print(f"     Colunas disponíveis: {cabecalho_excel}")
        print(f"     Continuando sem vincular dados da planilha...\n")
    else:
        for reg in registros_excel:
            id_val = str(reg.get(col_id_real, "") or "").strip().upper()
            if id_val:
                indice_excel.setdefault(id_val, []).append(reg)

        ids_duplicados = {k: v for k, v in indice_excel.items() if len(v) > 1}
        print(f"  ✅ {len(indice_excel)} IDs únicos no índice")

        if ids_duplicados:
            print(f"  ℹ️  {len(ids_duplicados)} ID(s) com múltiplos registros (correto — um KML, várias linhas):")
            for id_k in sorted(set(ids_duplicados.keys())):
                print(f"     • {id_k}: {len(ids_duplicados[id_k])} registros")

    print(f"\n📁 Buscando KMLs em: {KML_FOLDER}")
    if not os.path.exists(KML_FOLDER):
        print(f"  ❌ Pasta não encontrada: {KML_FOLDER}")
        return

    arquivos_kml = list(Path(KML_FOLDER).rglob("*.kml")) + list(Path(KML_FOLDER).rglob("*.KML"))
    arquivos_kml = sorted(set(arquivos_kml))
    print(f"  ✅ {len(arquivos_kml)} arquivos KML encontrados")

    print(f"\n🔗 Vinculando KMLs com a planilha...")
    print(f"   Chave: prefixo '{ID_REGEX}' do nome do arquivo KML → coluna '{COLUNA_ID}' do Excel\n")

    items               = []
    kml_vinculados      = 0
    kml_sem_vinculo     = 0
    kml_sem_poligono    = 0
    kml_sem_id          = 0
    nao_vinculados      = []
    ids_kml_processados = set()

    # ── Passo 1: agrupar todos os polígonos por ID ─────────────────
    # Cada MAP### pode ter vários arquivos KML (etapas, fases, etc.).
    # Todos os polígonos do mesmo ID são fundidos num único conjunto,
    # evitando que o VGV/unidades de cada linha Excel seja multiplicado
    # pelo número de arquivos KML do mesmo ID.
    kml_por_id = {}  # id_kml → { 'nome': str, 'poligonos': [] }

    for kml_path in arquivos_kml:
        nome_arquivo                       = kml_path.stem
        nome_kml_tag, poligonos, _centroide = extrair_coordenadas_kml(str(kml_path))

        if not poligonos:
            kml_sem_poligono += 1

        nome_display = nome_kml_tag or nome_arquivo

        id_kml = extrair_id(nome_arquivo)
        if not id_kml and nome_kml_tag:
            id_kml = extrair_id(nome_kml_tag)
        if not id_kml:
            kml_sem_id += 1

        if id_kml not in kml_por_id:
            kml_por_id[id_kml] = {'nome': nome_display, 'poligonos': []}

        kml_por_id[id_kml]['poligonos'].extend(poligonos)

    # ── Passo 2: criar itens — 1 item por linha do Excel ──────────
    # Para IDs com múltiplos KMLs: todos os polígonos ficam juntos
    # em cada item. O VGV/unidades de cada linha é contado apenas 1×.
    for id_kml, kml_info in kml_por_id.items():
        poligonos    = kml_info['poligonos']
        nome_display = kml_info['nome']

        # Centroide unificado de todos os polígonos deste ID
        todos_lats = [pt[0] for poly in poligonos for pt in poly]
        todos_lngs = [pt[1] for poly in poligonos for pt in poly]
        centroide  = (
            [sum(todos_lats) / len(todos_lats), sum(todos_lngs) / len(todos_lngs)]
            if todos_lats else None
        )

        registros_vinculados = indice_excel.get(id_kml, []) if id_kml else []

        if registros_vinculados:
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
        else:
            kml_sem_vinculo += 1
            nao_vinculados.append((nome_display, id_kml or "sem ID", f"ID={id_kml}"))
            items.append({
                "id": id_kml,
                "n":  nome_display,
                "p":  poligonos,
                "c":  centroide,
                "e":  None,
            })

    # ── Registros do Excel sem nenhum KML correspondente ──────────
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

    sem_localizacao = []
    for item in items:
        if not item["p"] and not item["c"]:
            e = item["e"] or {}
            sem_localizacao.append({
                "id":       item.get("id") or "—",
                "nome":     e.get("nome") or item["n"] or "—",
                "cidade":   e.get("cidade") or "—",
                "regional": e.get("regional") or "—",
                "motivo":   "sem KML" if item["e"] else "KML sem geometria",
            })

    # ── Estatísticas — fonte única: planilha Excel ─────────────────
    on_map = sum(1 for i in items if i["p"])

    def _soma_excel(col):
        """Soma os valores numéricos de uma coluna em todos os registros do Excel."""
        total = 0.0
        for reg in registros_excel:
            for k, v in reg.items():
                if k.strip().lower() == col.strip().lower():
                    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
                        total += v
                    break
        return total

    col_area   = COLUNAS.get("area_total")
    col_vgv    = COLUNAS.get("vgv_total")
    col_vgv_bt = COLUNAS.get("vgv_bt")
    col_units  = COLUNAS.get("total_unidades")

    total_area     = _soma_excel(col_area)   if col_area   else 0.0
    total_vgv      = _soma_excel(col_vgv)    if col_vgv    else 0.0
    total_vgv_bt   = _soma_excel(col_vgv_bt) if col_vgv_bt else 0.0
    total_unidades = _soma_excel(col_units)  if col_units  else 0.0

    col_on_off_real = COLUNAS.get("on_off")
    total_ativo   = sum(1 for r in registros_excel
                        for k, v in r.items()
                        if k.strip().lower() == (col_on_off_real or "").strip().lower() and v == 1)
    total_inativo = len(registros_excel) - total_ativo

    stats = {
        "total":          len(items),
        "total_planilha": len(registros_excel),
        "total_ativo":    total_ativo,
        "total_inativo":  total_inativo,
        "total_units":    round(total_unidades, 0),
        "total_area":     round(total_area, 2),
        "total_vgv":      round(total_vgv, 2),
        "total_vgv_bt":   round(total_vgv_bt, 2),
    }

    regional_summary = {}
    for item in items:
        if item["e"] and item["e"].get("regional"):
            r = item["e"]["regional"]
            if r not in regional_summary:
                regional_summary[r] = {"count": 0, "units": 0, "vgv": 0}
            regional_summary[r]["count"] += 1
            regional_summary[r]["units"] += item["e"].get("total_unidades") or 0
            regional_summary[r]["vgv"]   += item["e"].get("vgv_total") or 0

    try:
        last_updated = subprocess.check_output(
            ['git', 'log', '-1', '--format=%cd', '--date=format:%d/%m/%Y'],
            stderr=subprocess.DEVNULL
        ).decode().strip()
    except Exception:
        last_updated = datetime.now().strftime('%d/%m/%Y')

    data = {
        "items":            items,
        "colors":           CORES,
        "stats":            stats,
        "regional_summary": regional_summary,
        "last_updated":     last_updated,
    }

    js_content = "const DATA = " + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ";"
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(js_content)

    print(f"{'═'*60}")
    print(f"  ✅ data.js gerado com sucesso!")
    print(f"{'═'*60}")
    print(f"  📦 Total de itens gerados:                {len(items)}")
    print(f"  📋 Total de registros na planilha:        {stats['total_planilha']}")
    print(f"  🗺️  Com polígono KML:                      {on_map}")
    print(f"  📍 Sem localização:                       {len(sem_localizacao)}")
    print(f"  🔗 IDs KML vinculados à planilha:         {kml_vinculados}")
    print(f"  ❌ IDs KML sem vínculo na planilha:       {kml_sem_vinculo}")
    print(f"  📋 Registros só na planilha (sem KML):    {sem_kml}")
    print(f"  ⚠️  Arquivos KML sem geometria:            {kml_sem_poligono}")
    print(f"  🏷️  Arquivos KML sem ID reconhecível:      {kml_sem_id}")
    print(f"  💰 VGV Total:                          R$ {total_vgv:,.1f} mi")
    print(f"  🏘️  Total de unidades:                     {total_unidades:,.0f}")
    print(f"\n  📄 Arquivo gerado: {os.path.abspath(OUTPUT_PATH)}")

    if sem_localizacao:
        print(f"\n  📍 {len(sem_localizacao)} registro(s) SEM localização (sem polígono nem centroide):")
        col_id_w     = max(len(r["id"])       for r in sem_localizacao)
        col_nome_w   = max(len(r["nome"])     for r in sem_localizacao)
        col_cidade_w = max(len(r["cidade"])   for r in sem_localizacao)
        col_reg_w    = max(len(r["regional"]) for r in sem_localizacao)
        header = (
            f"     {'ID':<{col_id_w}}  {'Nome':<{col_nome_w}}  "
            f"{'Cidade':<{col_cidade_w}}  {'Regional':<{col_reg_w}}  Motivo"
        )
        print(header)
        print("     " + "─" * (len(header) - 5))
        for r in sorted(sem_localizacao, key=lambda x: (x["regional"], x["id"])):
            print(
                f"     {r['id']:<{col_id_w}}  {r['nome']:<{col_nome_w}}  "
                f"{r['cidade']:<{col_cidade_w}}  {r['regional']:<{col_reg_w}}  {r['motivo']}"
            )

    if nao_vinculados:
        print(f"\n  ⚠️  {len(nao_vinculados)} ID(s) KML SEM vínculo na planilha:")
        for nome, id_encontrado, path in nao_vinculados:
            label = f"ID={id_encontrado}" if id_encontrado != "sem ID" else "sem ID reconhecível"
            print(f"     • [{label}] \"{nome}\"")

    print(f"\n{'═'*60}")
    print("  👉 Próximo passo: faça commit e push do data.js para")
    print("     o repositório. O site atualizará automaticamente.")
    print(f"{'═'*60}\n")


if __name__ == "__main__":
    main()