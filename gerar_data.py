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
EXCEL_PATH = "AREAS_LAND_BANK.xlsx"

# Nome da aba/sheet da planilha (None = primeira aba)
EXCEL_SHEET = None

# Pasta onde estão os arquivos .kml (busca recursiva em subpastas)
KML_FOLDER = "kml"

# Arquivo de saída
OUTPUT_PATH = "data.js"

# Coluna da planilha usada para VINCULAR com a tag <name> do KML.
# O script compara o valor desta coluna com a tag <name> de dentro do arquivo KML.
# Se o KML não tiver <name>, usa o nome do arquivo como fallback.
COLUNA_CHAVE = "Nome"

# Mapeamento das colunas da planilha para os campos do sistema
# Formato: "campo_no_sistema": "nome_da_coluna_no_excel"
# Os nomes à direita devem ser EXATAMENTE iguais aos cabeçalhos da sua planilha.
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
    "NORTE":          "#c0392b",
    "NORDESTE I":     "#d35400",
    "NORDESTE II":    "#e67e22",
    "CENTRO OESTE":   "#27ae60",
    "CENTRO OESTE II":"#16a085",
    "SUDESTE":        "#2980b9",
    "SUL":            "#8e44ad",
    "TOCANTINS":      "#0e6655",
    "OESTE":          "#c2185b",
    "None":           "#7f8c8d",
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
    # Remove extensão .kml se houver
    texto = re.sub(r'\.KML$', '', texto)
    # Normaliza separadores de caminho — pega só o nome do arquivo
    texto = Path(texto).name
    return texto

def ler_excel(path, sheet=None):
    """Lê o Excel e retorna lista de dicionários, um por linha."""
    try:
        import openpyxl
    except ImportError:
        raise ImportError("Execute: pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb.active

    linhas = list(ws.iter_rows(values_only=True))
    if not linhas:
        return []

    cabecalho = [str(c).strip() if c is not None else "" for c in linhas[0]]
    registros = []

    for linha in linhas[1:]:
        if all(v is None for v in linha):
            continue  # pula linhas completamente vazias
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
      - poligonos: lista de polígonos [[lat,lng], ...]
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

    # ── Extrai <name> ──────────────────────────────────────────────
    # Prioridade: <Document><name>, depois primeiro <Placemark><name>,
    # depois qualquer <name> encontrado.
    nome_kml = None

    doc = root.find(".//Document")
    if doc is not None:
        tag = doc.find("name")
        if tag is not None and tag.text and tag.text.strip():
            nome_kml = tag.text.strip()

    if not nome_kml:
        # Tenta o primeiro Placemark
        pm = root.find(".//Placemark")
        if pm is not None:
            tag = pm.find("name")
            if tag is not None and tag.text and tag.text.strip():
                nome_kml = tag.text.strip()

    if not nome_kml:
        # Qualquer <name> no documento
        tag = root.find(".//name")
        if tag is not None and tag.text and tag.text.strip():
            nome_kml = tag.text.strip()

    # ── Extrai polígonos ───────────────────────────────────────────
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

    # Fallback: busca LineString também
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

    # Calcula centroide
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

    resultado = ler_excel(EXCEL_PATH, EXCEL_SHEET)
    registros_excel, cabecalho_excel = resultado

    # Monta índice: chave_normalizada → registro
    indice_excel = {}
    coluna_chave_real = None

    for col in cabecalho_excel:
        if col.strip().lower() == COLUNA_CHAVE.strip().lower():
            coluna_chave_real = col
            break

    if not coluna_chave_real:
        print(f"\n  ⚠️  Coluna '{COLUNA_CHAVE}' não encontrada no Excel.")
        print(f"     Colunas disponíveis: {cabecalho_excel}")
        print(f"     Ajuste COLUNA_CHAVE no início do script.")
        print(f"     Continuando sem vincular dados da planilha...\n")
    else:
        for reg in registros_excel:
            chave = normalizar(reg.get(coluna_chave_real, ""))
            if chave:
                indice_excel[chave] = reg

    # ── 2. Ler KMLs ───────────────────────────────────────────────
    print(f"\n📁 Buscando KMLs em: {KML_FOLDER}")
    if not os.path.exists(KML_FOLDER):
        print(f"  ❌ Pasta não encontrada: {KML_FOLDER}")
        print(f"     Verifique o caminho em KML_FOLDER no início do script.")
        return

    arquivos_kml = list(Path(KML_FOLDER).rglob("*.kml")) + list(Path(KML_FOLDER).rglob("*.KML"))
    arquivos_kml = sorted(set(arquivos_kml))
    print(f"  ✅ {len(arquivos_kml)} arquivos KML encontrados")

    # ── 3. Montar items ────────────────────────────────────────────
    print(f"\n🔗 Vinculando KMLs com a planilha...")
    print(f"   Chave: tag <name> do KML  →  coluna '{COLUNA_CHAVE}' da planilha")
    print(f"   (fallback para nome do arquivo caso <name> esteja vazia)\n")

    items = []
    vinculados = 0
    sem_poligono = 0
    usou_fallback_filename = 0
    nao_vinculados = []

    for kml_path in arquivos_kml:
        nome_arquivo = kml_path.stem  # nome do arquivo sem extensão (fallback)

        # Extrai <name>, geometria e centroide
        nome_kml_tag, poligonos, centroide = extrair_coordenadas_kml(str(kml_path))

        if not poligonos:
            sem_poligono += 1

        # Define a chave de vinculação:
        # 1ª prioridade: tag <name> do KML
        # 2ª prioridade: nome do arquivo (fallback)
        if nome_kml_tag:
            chave_kml    = normalizar(nome_kml_tag)
            nome_display = nome_kml_tag
            fonte_chave  = "<name>"
        else:
            chave_kml    = normalizar(nome_arquivo)
            nome_display = nome_arquivo
            fonte_chave  = "arquivo (sem <name>)"
            usou_fallback_filename += 1

        # Tenta vincular com Excel
        dados_excel = None
        if indice_excel:
            # Correspondência exata
            if chave_kml in indice_excel:
                dados_excel = indice_excel[chave_kml]
            else:
                # Correspondência parcial: um contém o outro
                for chave_excel, reg in indice_excel.items():
                    if chave_kml in chave_excel or chave_excel in chave_kml:
                        dados_excel = reg
                        break

        # Monta objeto e
        e = None
        if dados_excel:
            vinculados += 1
            e = {}
            for campo_sistema, col_excel in COLUNAS.items():
                if col_excel is None:
                    e[campo_sistema] = None
                    continue
                valor = None
                for k, v in dados_excel.items():
                    if k.strip().lower() == col_excel.strip().lower():
                        valor = serializar_valor(v)
                        break
                e[campo_sistema] = valor
        else:
            nao_vinculados.append((nome_display, fonte_chave, str(kml_path)))

        item = {
            "n": nome_display,
            "p": poligonos,
            "c": centroide,
            "e": e,
        }
        items.append(item)

    # ── 4. Montar itens sem KML (da planilha) ─────────────────────
    # IMPORTANTE: item["e"] usa as chaves do sistema (ex: "nome"), não os nomes
    # da coluna Excel (ex: "Nome"). Por isso usamos "nome" e não COLUNA_CHAVE aqui.
    chaves_kml_usadas = set()
    for item in items:
        if item["e"]:
            chave = normalizar(item["e"].get("nome", "") or "")
            chaves_kml_usadas.add(chave)

    sem_kml = 0
    for reg in registros_excel:
        chave = normalizar(reg.get(coluna_chave_real, "") if coluna_chave_real else "")
        if chave and chave not in chaves_kml_usadas:
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

            items.append({
                "n": reg.get(coluna_chave_real, "sem nome") if coluna_chave_real else "sem nome",
                "p": [],
                "c": None,
                "e": e,
            })
            sem_kml += 1

    # ── 5. Calcular estatísticas ───────────────────────────────────
    on_map         = sum(1 for i in items if i["p"])
    total_unidades = sum((i["e"].get("total_unidades") or 0) for i in items if i["e"])
    total_vgv      = sum((i["e"].get("vgv_total") or 0) for i in items if i["e"])
    total_vgv_bt   = sum((i["e"].get("vgv_bt") or 0) for i in items if i["e"])
    total_area     = sum(calcular_area(i["p"]) for i in items if i["p"])

    stats = {
        "total":       len(items),
        "on_map":      on_map,
        "total_units": total_unidades,
        "total_area":  round(total_area, 0),
        "total_vgv":   round(total_vgv, 2),
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

    # ── 6. Montar DATA e escrever data.js ─────────────────────────
    data = {
        "items":            items,
        "colors":           CORES,
        "stats":            stats,
        "regional_summary": regional_summary,
    }

    js_content = "const DATA = " + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ";"

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(js_content)

    # ── 7. Relatório final ────────────────────────────────────────
    print(f"{'═' * 60}")
    print(f"  ✅ data.js gerado com sucesso!")
    print(f"{'═' * 60}")
    print(f"  📦 Total de itens:               {len(items)}")
    print(f"  🗺️  Com polígono KML:              {on_map}")
    print(f"  🔗 Vinculados à planilha:        {vinculados}")
    print(f"  📋 Só na planilha (sem KML):     {sem_kml}")
    print(f"  ⚠️  KML sem coordenadas:          {sem_poligono}")
    print(f"  🏷️  Usou <name> como chave:       {len(arquivos_kml) - usou_fallback_filename}")
    print(f"  📄 Usou nome de arquivo (fallback):{usou_fallback_filename}")
    print(f"  💰 VGV Total:                    R$ {total_vgv:,.1f} mi")
    print(f"  🏘️  Total de unidades:             {total_unidades:,.0f}")
    print(f"\n  📄 Arquivo gerado: {os.path.abspath(OUTPUT_PATH)}")

    if nao_vinculados:
        print(f"\n  ⚠️  {len(nao_vinculados)} KML(s) SEM vínculo na planilha:")
        for nome, fonte, path in nao_vinculados:
            print(f"     • [{fonte}] \"{nome}\"")
            print(f"       → {path}")

    if usou_fallback_filename:
        print(f"\n  💡 {usou_fallback_filename} KML(s) sem tag <name> usaram o nome do arquivo.")
        print(f"     Considere adicionar <name> dentro do <Document> desses KMLs.")

    print(f"\n{'═' * 60}\n")
    print("  👉 Próximo passo: faça commit e push do data.js para")
    print("     o seu repositório. O site atualizará automaticamente.")
    print(f"{'═' * 60}\n")

if __name__ == "__main__":
    main()