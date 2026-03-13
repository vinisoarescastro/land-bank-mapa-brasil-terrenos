# 🗺️ Land Bank — Grupo Brasil Terrenos

Plataforma web interativa para visualização e gestão do banco de terrenos do **Grupo Brasil Terrenos**. A aplicação exibe empreendimentos imobiliários em um mapa interativo, com polígonos georreferenciados extraídos de arquivos KML e dados financeiros e operacionais provenientes de uma planilha Excel.

---

## ✨ Funcionalidades Principais

### Mapa Interativo
- Visualização de polígonos de terrenos sobre mapa base (OpenStreetMap) ou satélite (Esri/Google)
- Coloração dos polígonos por regional, com legenda dinâmica
- Marcadores de centroide para terrenos sem polígono cadastrado
- Popups com ficha completa do empreendimento ao clicar no mapa ou na lista lateral

### Sidebar de Navegação
- **Dashboard de estatísticas**: total de empreendimentos, terrenos no mapa (KML), total de unidades, área total e VGV
- **Busca textual** por nome do terreno, cidade ou regional
- **Filtros por regional** com chips coloridos (NORTE, NORDESTE I, NORDESTE II, CENTRO OESTE, SUDESTE, SUL, TOCANTINS, OESTE, etc.)
- **Filtro "Vinculados"**: exibe apenas terrenos com dados da planilha vinculados
- **Lista de terrenos** com nome, cidade, regional, unidades e indicadores visuais de status

### Popup de Detalhes
Ao selecionar um terreno, exibe:
- Tipo de empreendimento (Loteamento / Condomínio)
- Ano previsto de lançamento
- Área total (em hectares)
- Total de unidades
- VGV Total e VGV BT (em R$ milhões)
- Custo do terreno e custo de construção
- Participação Buriti (%)
- Status ON / OFF

### Responsividade Mobile
- Sidebar em modo drawer deslizante com overlay
- Botão de menu flutuante no mapa
- Suporte a toque e gestos

### Controle de Camadas
- Alternância entre mapa de ruas (OpenStreetMap) e imagem de satélite
- Painel colapsável de camadas no canto superior direito do mapa

---

## 🛠️ Tecnologias Utilizadas

### Frontend
| Tecnologia | Versão | Uso |
|---|---|---|
| HTML5 | — | Estrutura da página |
| CSS3 | — | Estilo e responsividade |
| JavaScript (Vanilla) | ES6+ | Lógica da aplicação |
| [Leaflet.js](https://leafletjs.com/) | 1.9.4 | Biblioteca de mapas interativos |
| DM Sans / JetBrains Mono | — | Tipografia (Google Fonts) |

### Pipeline de Dados (Python)
| Tecnologia | Uso |
|---|---|
| Python 3 | Script de geração do `data.js` |
| `openpyxl` | Leitura da planilha Excel (`.xlsx`) |
| `lxml` | Parsing dos arquivos KML |
| `json` / `math` / `pathlib` | Utilitários padrão |

### Formatos de Dados
| Formato | Uso |
|---|---|
| `.kml` | Polígonos georreferenciados dos terrenos (Google Earth) |
| `.xlsx` | Planilha com dados financeiros e operacionais |
| `data.js` | Arquivo gerado — alimenta o site em runtime |

---

## 📁 Estrutura do Repositório

```
mapa-teste/
│
├── index.html              # Página principal da aplicação
├── app.js                  # Lógica JavaScript (mapa, filtros, sidebar, popups)
├── styles.css              # Estilos CSS completos
│
├── data.js                 # GERADO: não editar manualmente
│                           # Gerado por gerar_data.py; contém todos os
│                           # terrenos, cores e estatísticas em formato JS
│
├── gerar_data.py           # Script Python para gerar o data.js
├── AREAS_LAND_BANK.xlsx    # Planilha com dados dos empreendimentos
│
├── kml/                    # Arquivos KML organizados por estado/cidade
│   ├── 10 PARÁ/
│   │   ├── PA - SANTARÉM/
│   │   │   ├── SANTARÉM.CJD06.kml
│   │   │   └── ...
│   │   └── PA - REDENÇÃO/
│   │       └── ...
│   └── ... (demais estados e regiões)
│
└── logo_brasil_terrenos.png  # Logotipo exibido na sidebar
```

---

## ⚙️ Pipeline de Atualização de Dados

O site é alimentado pelo arquivo `data.js`, gerado automaticamente pelo script `gerar_data.py`. O fluxo é:

```
AREAS_LAND_BANK.xlsx  ───┐
                         ├──► gerar_data.py  ──►  data.js  ──►  Site
kml/*.kml             ───┘
```

### Como gerar o `data.js`

**1. Instale as dependências Python:**
```bash
pip install openpyxl lxml
```

**2. Execute o script:**
```bash
python gerar_data.py
```

**3. Faça commit e push do `data.js` gerado:**
```bash
git add data.js
git commit -m "Atualiza data.js"
git push
```

O site atualizará automaticamente após o push.

### Configuração do script (`gerar_data.py`)

No início do arquivo, edite as seguintes variáveis conforme necessário:

| Variável | Padrão | Descrição |
|---|---|---|
| `EXCEL_PATH` | `AREAS_LAND_BANK.xlsx` | Caminho da planilha Excel |
| `EXCEL_SHEET` | `None` (primeira aba) | Nome da aba da planilha |
| `KML_FOLDER` | `kml` | Pasta com os arquivos KML |
| `OUTPUT_PATH` | `data.js` | Arquivo de saída |
| `COLUNA_CHAVE` | `Nome` | Coluna usada para vincular KML ↔ planilha |

### Lógica de vinculação KML ↔ Planilha

O script vincula cada arquivo KML com a linha correspondente na planilha pela tag `<n>` dentro do `<Document>` do KML. Caso a tag `<n>` não exista, utiliza o nome do arquivo como fallback. A correspondência é feita por texto normalizado (sem acentos, maiúsculas).

---

## 🏷️ Regionais e Cores

| Regional | Cor |
|---|---|
| NORTE | `#c0392b` |
| NORDESTE I | `#d35400` |
| NORDESTE II | `#e67e22` |
| CENTRO OESTE | `#27ae60` |
| CENTRO OESTE II | `#16a085` |
| SUDESTE | `#2980b9` |
| SUL | `#8e44ad` |
| TOCANTINS | `#0e6655` |
| OESTE | `#c2185b` |

---

## 🚀 Como Executar Localmente

Por ser uma aplicação estática (HTML + JS + CSS), basta servir os arquivos com qualquer servidor HTTP local:

```bash
# Com Python
python -m http.server 8000

# Com Node.js (npx)
npx serve .
```

Acesse em: `http://localhost:8000`

> **Atenção:** Não abra o `index.html` diretamente no navegador (`file://`) pois o carregamento do `data.js` pode ser bloqueado por restrições de CORS.

---

## 📊 Dados Exibidos no Dashboard

O dashboard na sidebar exibe as seguintes métricas agregadas, calculadas automaticamente pelo `gerar_data.py` a partir da planilha:

- **Empreendimentos**: total de itens no banco de terrenos
- **No Mapa (KML)**: quantidade de terrenos com polígono georreferenciado
- **Total de Unidades**: soma de unidades habitacionais de todos os empreendimentos
- **Área Total**: soma das áreas dos polígonos KML (em hectares)
- **VGV Total / VGV BT**: Valor Geral de Vendas total e da participação Buriti Terrenos (em R$ milhões)

---

## 🗂️ Estrutura do `data.js`

O arquivo gerado contém um único objeto JavaScript `DATA` com a seguinte estrutura:

```javascript
const DATA = {
  items: [
    {
      n: "Nome do Terreno",           // Nome (tag <n> do KML ou nome do arquivo)
      p: [[[lat, lng], ...]],         // Lista de polígonos
      c: [lat, lng],                  // Centroide
      e: {                            // Dados da planilha (null se sem vínculo)
        nome, codigo, regional, cidade,
        empreendimento, tipo, year,
        on_off, area_total, total_unidades,
        vgv_total, vgv_bt,
        custo_terreno, custo_construcao,
        participacao_buriti, data_lancamento
      }
    },
    // ...
  ],
  colors: { "NORTE": "#c0392b", ... },  // Mapeamento regional → cor
  stats: {                               // Estatísticas globais
    total, on_map, total_units,
    total_area, total_vgv, total_vgv_bt
  },
  regional_summary: {                    // Resumo por regional
    "NORTE": { count, units, vgv },
    // ...
  }
};
```