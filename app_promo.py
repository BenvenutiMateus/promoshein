import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
import difflib

# ================= FUNÇÕES =================

def ler_excel_promocao_com_formulas(file, sheet_name="PROMOÇÃO", header_row=0):
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet_name]
    df = pd.DataFrame(ws.values)

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df = df.loc[:, df.columns.notna()]
    df.columns = df.columns.astype(str)

    return df


def _normalize_merge_key(val):
    s = str(val)
    s = s.replace('.0', '')
    s = s.replace('MLB', '')
    s = s.replace('+T', '')
    s = s.replace("+T", "")
    s = s.replace("-1-", "-")
    s = s.replace("-2-", "-")
    s = s.replace("AB", "BC")
    s = s.replace("MH", "BC")
    s = s.replace("SAN", "BC")
    s = s.split()[0]  # Pega só a primeira parte antes de espaços
    s = s.strip()
    parts = [p.strip() for p in s.split('-') if p.strip()]
    lower = [p.lower() for p in parts]
    if any('kit' in p for p in lower):
        take = parts[:3]
    else:
        take = parts[:2]
    return '-'.join(take)


def _deduplicate_columns(df):
    df = df.copy()
    cols = list(df.columns)
    seen = {}
    new_cols = []
    for col in cols:
        base = str(col)
        candidate = base
        suffix = 1
        while candidate in seen:
            suffix += 1
            candidate = f"{base}_{suffix}"
        seen[candidate] = True
        new_cols.append(candidate)
    df.columns = new_cols
    return df, {}


def _suggest_column(columns, aliases):
    normalized_columns = [str(c).strip().lower() for c in columns]
    for alias in aliases:
        alias_norm = str(alias).strip().lower()
        for idx, col_name in enumerate(normalized_columns):
            if col_name == alias_norm:
                return columns[idx]
            if alias_norm in col_name:
                return columns[idx]
    return None


# ================= APP =================

st.set_page_config("Gerenciador de Promoções Shein", layout="wide")
st.title("📊 Gerenciador de Promoções para a Shein")

# ================= SIDEBAR =================

with st.sidebar:
    st.header("⚙️ Configuração")

    arquivo_skus = st.file_uploader("1️⃣ Planilha de SKUs / IDs", type=["xlsx", "xls", "csv"])
    arquivo_precos = st.file_uploader("2️⃣ Base de Preços", type=["xlsx", "xls", "csv"])

    if not arquivo_skus or not arquivo_precos:
        st.info("👆 Envie os dois arquivos")
        st.stop()

    # --- Leitura com seleção de sheet / linha de header ---
    st.sidebar.write("### Selecione sheet e linha de cabeçalho (quando aplicável)")

    # SKUs file
    df_skus = pd.read_excel(arquivo_skus)

    # Preços file
    xls_precos = pd.ExcelFile(arquivo_precos)
    # Ignora sempre a sheet 'hidden' (case-insensitive)
    sheet_names_precos = [s for s in xls_precos.sheet_names if s.strip().lower() != 'hidden']
    # default to PROMOÇÃO sheet when present (case-insensitive)
    def _find_default_sheet(names):
        for i, s in enumerate(names):
            if s.strip().lower() in ("promoção", "promocao"):
                return i
            if s.strip().lower() in ("promoção", "promocao"):
                return i
        return 0

    default_idx = next((i for i, s in enumerate(sheet_names_precos) if s.strip().lower() in ("promoção", "promocao")), 0)
    sheet_precos = st.selectbox("Sheet (Preços)", sheet_names_precos, index=default_idx, key="sheet_precos")
    header_precos = int(st.number_input("Linha do cabeçalho (Preços) - 1-based, 0=auto", min_value=0, max_value=50, value=0, key="header_precos"))
    header_row_precos = 0 if header_precos <= 1 else header_precos - 1

    try:
        df_precos = ler_excel_promocao_com_formulas(arquivo_precos, sheet_name=sheet_precos, header_row=header_row_precos)
    except Exception:
        df_precos = pd.read_excel(arquivo_precos, sheet_name=sheet_precos, header=header_row_precos)
    

    # Normaliza nomes de colunas (remove espaços invisíveis)
    df_skus.columns = df_skus.columns.astype(str).str.strip()
    df_precos.columns = df_precos.columns.astype(str).str.strip()
    
        # Evita erro no Streamlit/PyArrow com nomes de coluna duplicados
    df_skus, skus_renamed = _deduplicate_columns(df_skus)
    df_precos, precos_renamed = _deduplicate_columns(df_precos)

    # Inicializa variáveis de detecção para evitar NameError
    skc_col = None

    # Limpeza base de preços
    colunas_remover = [
        "descrição", "descricao", "valor a receber", "peso",
        "frete", "taxa", "redução", "reducao", "bruto", "publicação"
    ]

    df_precos = df_precos.loc[:, [c for c in df_precos.columns if not any(r in c.lower() for r in colunas_remover)]]


    df_skus = df_skus.loc[:, [c for c in df_skus.columns if not any(m in c.lower() for m in ['shein'])]]

    st.success("✅ Arquivos carregados")

    st.divider()


    # --- Mapeamento de colunas (auto-sugestão básica) ---
    st.sidebar.write("### Mapeamento de colunas (ajuste se necessário)")

    # sugestões para SKUs
    skc_aliases = ["skc"]
    num_item_aliases = ["número do item", "numero do item", "item number", "item no", "numero item"]
    sku_aliases = ["Número do item", "numero do item", "item number", "item no", "numero item", "sku", "id", "id sku"]

    # Recalcula sugestões sempre após leitura do df_skus
    def _auto_suggest(df, aliases):
        # Tenta por nome
        col = _suggest_column(df.columns, aliases)
        if col:
            return col
        # Tenta por conteúdo: procura coluna com valores que parecem SKU (alfanumérico, tamanho típico)
        for c in df.columns:
            vals = df[c].astype(str).str.upper().str.strip().dropna().unique()
            if any(len(v) >= 6 and v.isalnum() for v in vals[:10]):
                return c
        return None

    sugestao_skc = _auto_suggest(df_skus, skc_aliases)
    sugestao_sku = _auto_suggest(df_skus, sku_aliases)


    col_skc_sel = 'SKC'

    col_match_skus = st.selectbox(
        "Coluna de match (SKUs)",
        list(df_skus.columns),
        index=(list(df_skus.columns).index(sugestao_sku) if sugestao_sku in df_skus.columns else 0),
        key="col_match_skus_map"
    )

    # sugestões para preços

    price_aliases = ["preço", "preco", "valor", "price"]
    sugestao_price = _suggest_column(df_precos.columns, price_aliases)
    sugestao_precos_match = _suggest_column(df_precos.columns, sku_aliases + num_item_aliases)

    colunas_match_precos = [c for c in df_precos.columns if not any(m in c.lower() for m in ['shein'])]
    col_match_precos = st.selectbox(
        "Coluna de match (Preços)",
        colunas_match_precos,
        index=(colunas_match_precos.index(sugestao_precos_match) if sugestao_precos_match in colunas_match_precos else 0),
        key="col_match_precos_map"
    )

    # Filtra colunas de preço relacionadas ao marketplace selecionado
    def _colunas_preco_marketplace(cols):
        return [c for c in cols if 'shein' in c.lower()]

    colunas_preco_filtradas = _colunas_preco_marketplace(df_precos.columns)
    # Garante que pelo menos uma coluna apareça
    if not colunas_preco_filtradas:
        colunas_preco_filtradas = list(df_precos.columns)

    col_preco = st.selectbox(
        "Coluna de Preço",
        colunas_preco_filtradas,
        index=(colunas_preco_filtradas.index(sugestao_price) if sugestao_price in colunas_preco_filtradas else 0),
        key="col_preco_map"
    )

    # Exibe pré-visualizações rápidas para ajudar a confirmar mapeamento
    st.sidebar.write("Preview SKUs:")
    st.sidebar.dataframe(df_skus.head(5))
    st.sidebar.write("Preview Preços:")
    st.sidebar.dataframe(df_precos.head(5))

# ================= PROCESSAMENTO =================

# 🔒 Coluna canônica de ID (NUNCA some)
if "SKU" in df_skus.columns and "SKU_SHEIN" not in df_skus.columns:
    df_skus = df_skus.rename(columns={"SKU": "SKU_SHEIN"})

if col_match_skus == "SKU":
    col_match_skus = "SKU_SHEIN"

df_skus["ID_BASE"] = df_skus[col_match_skus]

df_skus["_MERGE_KEY"] = (
    df_skus[col_match_skus]
    .astype(str)
).apply(_normalize_merge_key)



# Ajuste especial para Shein: prefira a coluna "Número do item" como chave de
# agrupamento (isso agrupa variações pelo SKU pai). Também detectamos a
# coluna `SKC` para uso no export quando disponível.

def _find_skc_col(df):
    for c in df.columns:
        if c.strip().lower() == "skc" or "skc" in c.lower():
            return c
    return None

skc_col = _find_skc_col(df_skus)



try:
    if 'col_skc_sel' in locals() and col_skc_sel and col_skc_sel in df_skus.columns:
        skc_col = col_skc_sel
except Exception:
    pass

# Criar _MERGE_KEY em df_precos a partir da coluna selecionada
df_precos["_MERGE_KEY"] = (
    df_precos[col_match_precos]
    .astype(str)
    .str.replace(r"\\.0$", "", regex=True)
)

# Permite múltiplos IDs separados por vírgula
df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].str.split(",")

# Explode IDs múltiplos
df_precos = df_precos.explode("_MERGE_KEY")
df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].astype(str).str.replace(".0","", regex=False).str.strip()
df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].apply(_normalize_merge_key)


# Opção: colapsar linhas idênticas por SKC (útil para Shein)
collapse_skc = False
collapse_post_merge = False

st.sidebar.write(f"Detectado SKC: {skc_col}")
collapse_skc = st.sidebar.checkbox("Colapsar linhas idênticas por SKC", value=True)

if collapse_skc:
    # Se houver SKC em df_precos, colapsa antes do merge; se houver apenas em
    # df_skus, agendamos o colapso para depois do merge.
    if skc_col and skc_col in df_precos.columns:
        before = len(df_precos)

        def _first_nonnull(s):
            s2 = s.dropna()
            return s2.iloc[0] if not s2.empty else s.iloc[0]

        agg_map = {c: 'first' for c in df_precos.columns}
        if col_preco in df_precos.columns:
            agg_map[col_preco] = _first_nonnull
        agg_map['_MERGE_KEY'] = 'first'

        df_precos = df_precos.groupby(skc_col, as_index=False).agg(agg_map)
        after = len(df_precos)
        removed = before - after
        if removed > 0:
            st.sidebar.info(f"✅ Colapsadas {removed} linhas idênticas por SKC — agora {after} SKC únicos.")
    elif skc_col and skc_col in df_skus.columns:
        collapse_post_merge = True
        st.sidebar.info("ℹ️ SKC detectado apenas em SKUs — colapso será feito após o merge.")
    else:
        st.sidebar.warning("⚠️ SKC não detectado em SKUs nem em Preços — não foi possível colapsar por SKC.")

# Remove colisões (preserva ID_BASE)
colisoes = set(df_skus.columns) & set(df_precos.columns)
colisoes.discard("_MERGE_KEY")
colisoes.discard("ID_BASE")

df_skus_limpo = df_skus.drop(columns=list(colisoes))

# Merge
df_merged = df_skus_limpo.merge(df_precos, on="_MERGE_KEY", how="left")
df_merged.drop(columns="_MERGE_KEY", inplace=True)

# Se foi marcado para colapsar por SKC somente após o merge (quando SKC
# existe apenas no arquivo da Shein), aplicamos o agrupamento aqui.
if 'collapse_post_merge' in globals() and collapse_post_merge:
    if skc_col and skc_col in df_merged.columns:
        before = len(df_merged)

        def _first_nonnull(s):
            s2 = s.dropna()
            return s2.iloc[0] if not s2.empty else s.iloc[0]

        agg_map = {c: 'first' for c in df_merged.columns}
        if col_preco in df_merged.columns:
            agg_map[col_preco] = _first_nonnull

        df_merged = df_merged.groupby(skc_col, as_index=False).agg(agg_map)
        after = len(df_merged)
        removed = before - after
        if removed > 0:
            st.sidebar.info(f"✅ Colapsadas {removed} linhas idênticas por SKC após o merge — agora {after} SKC únicos.")

# ================= TABS =================

tab1, tab2, tab3 = st.tabs(["📋 Dados", "🔗 Match", "⬇️ Download"])

# ---------- TAB 1 ----------
with tab1:
    st.subheader("SKUs")
    st.dataframe(df_skus, use_container_width=True)

    st.subheader("Preços")
    st.dataframe(df_precos, use_container_width=True)

# ---------- TAB 2 ----------
with tab2:
    st.subheader("🔗 Resultado do Match")

    total = len(df_merged)
    matched = df_merged[col_preco].notna().sum()
    nao_matched = df_merged[col_preco].isna().sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total SKUs", total)
    c2.metric("Matched", matched)
    c3.metric("Não encontrados", nao_matched)

    st.divider()

    st.write("### 📌 Amostra geral")
    st.dataframe(df_merged, use_container_width=True)

    st.divider()

    df_nao_encontrados = df_merged[df_merged[col_preco].isna()]

    if not df_nao_encontrados.empty:
        st.warning(f"⚠️ {len(df_nao_encontrados)} SKUs não tiveram match")

        if st.checkbox("🔍 Mostrar apenas SKUs não encontrados"):
            st.dataframe(
                df_nao_encontrados,
                use_container_width=True
            )
    else:
        st.success("🎉 Todos os SKUs tiveram match!")

# ---------- TAB 3 ----------
with tab3:
    df_final = df_merged.copy()

    # Trata #REF!, texto, etc
    df_final[col_preco] = pd.to_numeric(df_final[col_preco], errors="coerce")
    # Preenche valores não encontrados com -1
    df_final[col_preco] = df_final[col_preco].fillna(-1)
    df_final[col_preco] = df_final[col_preco].round(2)

    # Para Shein: se detectamos a coluna `SKC`, exportamos `SKC` + preço de campanha
    sku_export_col = "SKU_SHEIN" if "SKU_SHEIN" in df_final.columns else "SKU" if "SKU" in df_final.columns else None

    if skc_col and skc_col in df_final.columns:
        df_export = df_final[[skc_col, sku_export_col, col_preco]].copy()
        df_export = df_export.rename(columns={
            skc_col: "SKC (obrigatório)",
            sku_export_col: "SKU_SHEIN",
            col_preco: f"Preço de campanha"
        })
    else:
        df_export = df_final[["ID_BASE", sku_export_col, col_preco]].copy()

        df_export = df_export.rename(columns={
            "ID_BASE": "ID",
            sku_export_col: "SKU_SHEIN",
            col_preco: f"Preço de campanha"
        })

    st.info(f"📊 {len(df_export)} registros prontos")

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False)

    st.download_button(
        "📥 Baixar Excel",
        buffer.getvalue(),
        file_name=f"promo_shein_{datetime.now():%d%m%Y_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
