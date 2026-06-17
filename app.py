from dataclasses import replace
from datetime import date, datetime
from typing import Dict, Set, Tuple

import pandas as pd
import streamlit as st

from processor import (
    ProcessConfig,
    processar_precos,
    build_report_excel,
    build_protheus_csv,
    build_log_txt,
    CommonPricingConfig,
    processar_precificacao_comum,
    build_common_report_excel,
    build_common_protheus_csv,
    build_common_log_txt,
    build_zip_generic,
    normalize_table_code,
    normalize_sku,
)

st.set_page_config(page_title="GERADOR DE TABELAS DE PREÇOS EM MASSA", layout="wide")

st.title("GERADOR DE TABELAS DE PREÇOS EM MASSA")
st.caption(
    "Home com dois módulos: Precificação Comum e Adequação Mercado Livre. "
    "Versão V12 - seleção de SKUs para exportação antes do download."
)

if "resultado_ml" not in st.session_state:
    st.session_state["resultado_ml"] = None
if "resultado_comum" not in st.session_state:
    st.session_state["resultado_comum"] = None

st.sidebar.header("Home")
modulo = st.sidebar.radio(
    "Escolha a função",
    ["Precificação Comum", "Adequação ML"],
    index=0,
)


def _normalizar_sku_exportacao(value) -> str:
    sku = normalize_sku(value)
    return str(sku or "").strip()


def _filtrar_df_por_skus(df: pd.DataFrame, selected_skus: Set[str]) -> pd.DataFrame:
    """Filtra DataFrames de saída Protheus sem alterar a lógica do motor."""
    if df is None or df.empty or "SKU" not in df.columns:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    filtrado = df.copy()
    filtrado["_SKU_FILTRAGEM_EXPORTACAO"] = filtrado["SKU"].apply(_normalizar_sku_exportacao)
    filtrado = filtrado[filtrado["_SKU_FILTRAGEM_EXPORTACAO"].isin(selected_skus)].drop(columns=["_SKU_FILTRAGEM_EXPORTACAO"])
    return filtrado.reset_index(drop=True)


def _analise_com_coluna_exportar(df: pd.DataFrame, selected_skus: Set[str]) -> pd.DataFrame:
    if df is None or df.empty or "SKU" not in df.columns:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    analise = df.copy()
    skus_norm = analise["SKU"].apply(_normalizar_sku_exportacao)
    analise.insert(0, "Exportar?", skus_norm.isin(selected_skus))
    return analise


def _render_editor_exportacao(df: pd.DataFrame, key: str, default_skus: Set[str]) -> Tuple[Set[str], pd.DataFrame]:
    """Mostra a revisão de SKUs com checkbox e retorna os SKUs marcados."""
    if df is None or df.empty or "SKU" not in df.columns:
        st.warning("Não há SKUs disponíveis para revisão de exportação.")
        return set(), df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    tabela = df.copy()
    skus_norm = tabela["SKU"].apply(_normalizar_sku_exportacao)
    tabela.insert(0, "Exportar?", skus_norm.isin(default_skus))

    disabled_cols = [col for col in tabela.columns if col != "Exportar?"]
    edited = st.data_editor(
        tabela,
        use_container_width=True,
        hide_index=True,
        disabled=disabled_cols,
        column_config={
            "Exportar?": st.column_config.CheckboxColumn(
                "Exportar?",
                help="Desmarque para retirar o SKU dos arquivos Protheus desta execução.",
                default=True,
            )
        },
        key=key,
    )

    selected = set(
        edited.loc[edited["Exportar?"] == True, "SKU"]  # noqa: E712
        .apply(_normalizar_sku_exportacao)
        .dropna()
        .astype(str)
        .tolist()
    )
    return selected, edited


def _build_ml_filtered_result(result, selected_skus: Set[str]):
    analise = _analise_com_coluna_exportar(result.analise, selected_skus)
    protheus = _filtrar_df_por_skus(result.protheus, selected_skus)
    protheus_013 = _filtrar_df_por_skus(result.protheus_013, selected_skus)
    return replace(result, analise=analise, protheus=protheus, protheus_013=protheus_013)


def _build_common_filtered_result(result, selected_skus: Set[str]):
    analise = _analise_com_coluna_exportar(result.analise, selected_skus)
    saidas = {
        tabela: _filtrar_df_por_skus(df, selected_skus)
        for tabela, df in result.saidas.items()
    }
    return replace(result, analise=analise, saidas=saidas)


def _build_ml_outputs(payload, filtered_result):
    cfg = payload["config"]
    data_nome = payload["data_nome"]
    selecionados = payload.get("selecionados", {})
    outputs = {}

    if selecionados.get("relatorio"):
        report_name = f"relatorio_tabela_007_013_ml_{data_nome}.xlsx"
        outputs[report_name] = {
            "label": "Relatório completo",
            "bytes": build_report_excel(filtered_result),
            "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }

    if selecionados.get("tabela_007"):
        tabela_007_norm = normalize_table_code(cfg.tabela_codigo)
        fname = f"protheus_tabela_{tabela_007_norm}_{data_nome}.csv"
        outputs[fname] = {
            "label": f"Protheus {tabela_007_norm}",
            "bytes": build_protheus_csv(filtered_result, cfg, tabela=cfg.tabela_codigo),
            "mime": "text/csv",
        }

    if selecionados.get("tabela_013"):
        tabela_013_norm = normalize_table_code(cfg.tabela_codigo_013)
        fname = f"protheus_tabela_{tabela_013_norm}_{data_nome}.csv"
        outputs[fname] = {
            "label": f"Protheus {tabela_013_norm}",
            "bytes": build_protheus_csv(filtered_result, cfg, tabela=cfg.tabela_codigo_013),
            "mime": "text/csv",
        }

    if selecionados.get("log"):
        log_name = f"log_tabela_007_013_ml_{data_nome}.txt"
        outputs[log_name] = {"label": "Log TXT", "bytes": build_log_txt(filtered_result), "mime": "text/plain"}

    zip_name = f"pacote_tabela_007_013_ml_{data_nome}.zip"
    zip_bytes = build_zip_generic({name: item["bytes"] for name, item in outputs.items()})
    return outputs, zip_name, zip_bytes


def _build_common_outputs(payload, filtered_result):
    data_nome = payload["data_nome"]
    selecionados = payload.get("selecionados", {})
    outputs = {}

    if selecionados.get("relatorio"):
        report_name = f"relatorio_precificacao_comum_{data_nome}.xlsx"
        outputs[report_name] = {
            "label": "Relatório",
            "bytes": build_common_report_excel(filtered_result),
            "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }

    if selecionados.get("log"):
        log_name = f"log_precificacao_comum_{data_nome}.txt"
        outputs[log_name] = {"label": "Log TXT", "bytes": build_common_log_txt(filtered_result), "mime": "text/plain"}

    for tabela in selecionados.get("tabelas", []):
        tabela_norm = normalize_table_code(tabela)
        fname = f"protheus_tabela_{tabela_norm}_{data_nome}.csv"
        outputs[fname] = {
            "label": f"Protheus {tabela_norm}",
            "bytes": build_common_protheus_csv(filtered_result, tabela_norm),
            "mime": "text/csv",
        }

    zip_name = f"pacote_precificacao_comum_{data_nome}.zip"
    zip_bytes = build_zip_generic({name: item["bytes"] for name, item in outputs.items()})
    return outputs, zip_name, zip_bytes


def _render_downloads(outputs: Dict[str, Dict], zip_name: str, zip_bytes: bytes, key_prefix: str):
    download_items = list(outputs.items()) + [(zip_name, {"label": "ZIP selecionados", "bytes": zip_bytes, "mime": "application/zip"})]
    cols = st.columns(min(len(download_items), 6))
    for idx, (fname, item) in enumerate(download_items):
        with cols[idx % len(cols)]:
            st.download_button(
                item["label"],
                item["bytes"],
                fname,
                mime=item["mime"],
                key=f"{key_prefix}_down_{idx}_{fname}",
            )


def render_adequacao_ml():
    st.header("Adequação ML")
    st.caption(
        "Atualiza a tabela 007 com base nos preços revisados do Mercado Livre, gera 007, 013, relatório e log."
    )

    with st.sidebar:
        st.divider()
        st.subheader("Parâmetros - Adequação ML")
        codigo_tabela = st.text_input("Código da tabela Protheus base", value="007", max_chars=3, key="ml_codigo")
        codigo_tabela_013 = st.text_input("Código da tabela Protheus +12%", value="013", max_chars=3, key="ml_codigo_013")
        acrescimo_013 = st.number_input(
            "Acréscimo da tabela 013 sobre a 007",
            min_value=0.0,
            max_value=1.0,
            value=0.12,
            step=0.01,
            format="%.2f",
            key="ml_acrescimo_013",
        )
        desconto_ml = st.number_input(
            "Desconto sobre Preço ML",
            min_value=0.0,
            max_value=0.9,
            value=0.10,
            step=0.01,
            format="%.2f",
            key="ml_desconto",
        )
        markup_min = st.number_input("Alerta: markup menor que", min_value=0.0, max_value=10.0, value=1.20, step=0.05, format="%.2f", key="ml_markup")
        margem_critica = st.number_input("Alerta crítico: margem menor que", min_value=0.0, max_value=1.0, value=0.35, step=0.01, format="%.2f", key="ml_margem_critica")
        margem_atencao = st.number_input("Alerta atenção: margem até", min_value=0.0, max_value=1.0, value=0.50, step=0.01, format="%.2f", key="ml_margem_atencao")
        data_alt = st.date_input("Data da alteração Protheus", value=date.today(), format="DD/MM/YYYY", key="ml_data")
        dup_rule = st.selectbox("Se houver SKU duplicado no ML, usar", ["ultima", "primeira"], index=0, key="ml_dup")

    st.subheader("1. Envie os arquivos")
    col1, col2, col3 = st.columns(3)
    with col1:
        tabela_007_file = st.file_uploader("Tabela 007 oficial (.xlsx)", type=["xlsx"], key="ml_tabela007")
    with col2:
        lista_ml_file = st.file_uploader("PREÇOS ML - REVISADO (.xlsx)", type=["xlsx"], key="ml_lista")
    with col3:
        anterior_file = st.file_uploader("Relatório anterior gerado pelo app (opcional)", type=["xlsx"], key="ml_anterior")

    st.info(
        "A lista ML deve conter as colunas A=SKU, B=Preço ML e C=Custo Médio. "
        "A tabela 007 oficial deve seguir a estrutura já usada na rotina ML."
    )

    st.subheader("2. Escolha o que deseja gerar")
    opt1, opt2, opt3, opt4 = st.columns(4)
    with opt1:
        gerar_relatorio = st.checkbox("Relatório completo (.xlsx)", value=True, key="ml_opt_relatorio")
    with opt2:
        gerar_007 = st.checkbox(f"Tabela {str(codigo_tabela).zfill(3)} (.csv)", value=True, key="ml_opt_007")
    with opt3:
        gerar_013 = st.checkbox(f"Tabela {str(codigo_tabela_013).zfill(3)} (.csv)", value=True, key="ml_opt_013")
    with opt4:
        gerar_log = st.checkbox("Log TXT", value=True, key="ml_opt_log")

    alguma_saida = any([gerar_relatorio, gerar_007, gerar_013, gerar_log])

    btn_col1, btn_col2 = st.columns([1, 4])
    with btn_col1:
        gerar = st.button(
            "Gerar ML",
            type="primary",
            disabled=not (tabela_007_file and lista_ml_file and alguma_saida),
            key="ml_gerar",
        )
    with btn_col2:
        limpar = st.button("Limpar resultado ML", disabled=st.session_state["resultado_ml"] is None, key="ml_limpar")

    if not alguma_saida:
        st.warning("Selecione pelo menos uma saída para gerar.")

    if limpar:
        st.session_state["resultado_ml"] = None
        st.success("Resultado ML limpo.")

    if gerar:
        cfg = ProcessConfig(
            tabela_codigo=str(codigo_tabela).zfill(3),
            tabela_codigo_013=str(codigo_tabela_013).zfill(3),
            acrescimo_tabela_013=float(acrescimo_013),
            desconto_ml=float(desconto_ml),
            markup_minimo=float(markup_min),
            margem_critica=float(margem_critica),
            margem_atencao=float(margem_atencao),
            usar_ocorrencia_duplicada=dup_rule,
        )
        with st.spinner("Processando Adequação ML..."):
            result = processar_precos(
                tabela_007_file=tabela_007_file,
                lista_ml_file=lista_ml_file,
                relatorio_anterior_file=anterior_file,
                config=cfg,
                data_alteracao=data_alt,
            )
            data_nome = data_alt.strftime("%Y%m%d")

        st.session_state["resultado_ml"] = {
            "result": result,
            "config": cfg,
            "data_nome": data_nome,
            "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "generation_id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "selecionados": {
                "relatorio": gerar_relatorio,
                "tabela_007": gerar_007,
                "tabela_013": gerar_013,
                "log": gerar_log,
            },
        }

    payload = st.session_state.get("resultado_ml")
    if payload:
        result = payload["result"]
        st.success(f"Processamento ML gerado com sucesso. Resultado mantido desde {payload['gerado_em']}.")
        r = result.resumo_dict
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("SKUs ML", r.get("SKUs na Lista ML revisada", 0))
        m2.metric("Encontrados 007", r.get("SKUs encontrados na tabela 007", 0))
        m3.metric("Preços alterados", r.get("Preços alterados", 0))
        m4.metric("Prejuízo", r.get("Itens com prejuízo vs custo médio", 0))
        m5.metric("Alertas", len(result.alertas))

        st.subheader("3. Revisar SKUs para exportação")
        st.caption(
            "Todos os SKUs vêm marcados por padrão. Desmarque os que não devem entrar nos CSVs Protheus desta execução. "
            "O processamento continua visível para auditoria."
        )
        default_skus = set(result.protheus["SKU"].astype(str).apply(_normalizar_sku_exportacao).tolist()) if not result.protheus.empty else set()
        selected_skus, selection_table = _render_editor_exportacao(result.analise, f"ml_editor_{payload['generation_id']}", default_skus)
        filtered_result = _build_ml_filtered_result(result, selected_skus)

        c1, c2, c3 = st.columns(3)
        c1.metric("SKUs selecionados", len(selected_skus))
        c2.metric("Linhas CSV 007", len(filtered_result.protheus))
        c3.metric("Linhas CSV 013", len(filtered_result.protheus_013))

        outputs, zip_name, zip_bytes = _build_ml_outputs(payload, filtered_result)

        st.subheader("4. Baixar arquivos selecionados")
        _render_downloads(outputs, zip_name, zip_bytes, "ml")

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Análise", "Alertas", "Comparativo", "Log", "Protheus 007", "Protheus 013"])
        with tab1:
            st.dataframe(filtered_result.analise, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(result.alertas, use_container_width=True, hide_index=True)
        with tab3:
            st.dataframe(result.comparativo, use_container_width=True, hide_index=True) if not result.comparativo.empty else st.write("Nenhum comparativo disponível.")
        with tab4:
            st.dataframe(result.log, use_container_width=True, hide_index=True)
            st.text_area("Prévia do log TXT", result.log_text, height=280, key="ml_log_preview")
        with tab5:
            if payload.get("selecionados", {}).get("tabela_007"):
                st.dataframe(filtered_result.protheus, use_container_width=True, hide_index=True)
            else:
                st.write("Tabela 007 não foi selecionada para exportação nesta execução.")
        with tab6:
            if payload.get("selecionados", {}).get("tabela_013"):
                st.dataframe(filtered_result.protheus_013, use_container_width=True, hide_index=True)
            else:
                st.write("Tabela 013 não foi selecionada para exportação nesta execução.")


def render_precificacao_comum():
    st.header("Precificação Comum")
    st.caption(
        "Gera as tabelas 001, 004, 012, 007 e 013 a partir de uma planilha com A=SKU, B=Custo Médio e C=Preço 001."
    )

    with st.sidebar:
        st.divider()
        st.subheader("Parâmetros - Precificação Comum")
        desconto_007 = st.number_input(
            "Desconto da tabela 007 sobre a 001",
            min_value=0.0,
            max_value=0.9,
            value=0.16,
            step=0.01,
            format="%.2f",
            key="comum_desconto_007",
        )
        acrescimo_013 = st.number_input(
            "Acréscimo da tabela 013 sobre a 007",
            min_value=0.0,
            max_value=1.0,
            value=0.12,
            step=0.01,
            format="%.2f",
            key="comum_acrescimo_013",
        )
        data_alt = st.date_input("Data da alteração Protheus", value=date.today(), format="DD/MM/YYYY", key="comum_data")
        dup_rule = st.selectbox("Se houver SKU duplicado na base, usar", ["ultima", "primeira"], index=0, key="comum_dup")

    st.subheader("1. Envie a planilha base")
    base_file = st.file_uploader("Planilha de Precificação Comum (.xlsx)", type=["xlsx"], key="comum_base")
    st.info(
        "Estrutura esperada: coluna A = SKU, coluna B = Custo Médio, coluna C = Preço 001. "
        "O app também tenta reconhecer os cabeçalhos se a planilha tiver nomes de colunas."
    )

    st.subheader("2. Escolha o que deseja gerar")
    c1, c2, c3, c4 = st.columns(4)
    c5, c6, c7 = st.columns(3)
    with c1:
        gerar_relatorio = st.checkbox("Relatório completo (.xlsx)", value=True, key="comum_opt_relatorio")
    with c2:
        gerar_001 = st.checkbox("Tabela 001 (.csv)", value=True, key="comum_opt_001")
    with c3:
        gerar_004 = st.checkbox("Tabela 004 (.csv)", value=True, key="comum_opt_004")
    with c4:
        gerar_012 = st.checkbox("Tabela 012 (.csv)", value=True, key="comum_opt_012")
    with c5:
        gerar_007 = st.checkbox("Tabela 007 (.csv)", value=True, key="comum_opt_007")
    with c6:
        gerar_013 = st.checkbox("Tabela 013 (.csv)", value=True, key="comum_opt_013")
    with c7:
        gerar_log = st.checkbox("Log TXT", value=True, key="comum_opt_log")

    tabelas_selecionadas = [
        tabela
        for tabela, selecionada in {
            "001": gerar_001,
            "004": gerar_004,
            "012": gerar_012,
            "007": gerar_007,
            "013": gerar_013,
        }.items()
        if selecionada
    ]
    alguma_saida = any([gerar_relatorio, gerar_log, bool(tabelas_selecionadas)])

    btn_col1, btn_col2 = st.columns([1, 4])
    with btn_col1:
        gerar = st.button(
            "Gerar Precificação",
            type="primary",
            disabled=not (base_file and alguma_saida),
            key="comum_gerar",
        )
    with btn_col2:
        limpar = st.button("Limpar resultado comum", disabled=st.session_state["resultado_comum"] is None, key="comum_limpar")

    if not alguma_saida:
        st.warning("Selecione pelo menos uma saída para gerar.")

    if limpar:
        st.session_state["resultado_comum"] = None
        st.success("Resultado de Precificação Comum limpo.")

    if gerar:
        cfg = CommonPricingConfig(
            desconto_007_sobre_001=float(desconto_007),
            acrescimo_013_sobre_007=float(acrescimo_013),
            usar_ocorrencia_duplicada=dup_rule,
        )
        with st.spinner("Processando Precificação Comum..."):
            result = processar_precificacao_comum(base_file, config=cfg, data_alteracao=data_alt)
            data_nome = data_alt.strftime("%Y%m%d")

        st.session_state["resultado_comum"] = {
            "result": result,
            "config": cfg,
            "data_nome": data_nome,
            "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "generation_id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "selecionados": {
                "relatorio": gerar_relatorio,
                "log": gerar_log,
                "tabelas": tabelas_selecionadas,
            },
        }

    payload = st.session_state.get("resultado_comum")
    if payload:
        result = payload["result"]
        st.success(f"Processamento de Precificação Comum gerado com sucesso. Resultado mantido desde {payload['gerado_em']}.")
        r = result.resumo_dict
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("SKUs base", r.get("SKUs válidos na base", 0))
        m2.metric("Tabela 001", r.get("Linhas geradas tabela 001", 0))
        m3.metric("Tabela 004/012", r.get("Linhas geradas tabela 004", 0))
        m4.metric("Tabela 007/013", r.get("Linhas geradas tabela 007", 0))
        m5.metric("Intercorrências", len(result.log))

        st.subheader("3. Revisar SKUs para exportação")
        st.caption(
            "Todos os SKUs vêm marcados por padrão. Desmarque os que não devem entrar nos CSVs Protheus desta execução. "
            "A análise permanece disponível para auditoria."
        )
        default_skus = set()
        for df_saida in result.saidas.values():
            if df_saida is not None and not df_saida.empty and "SKU" in df_saida.columns:
                default_skus.update(df_saida["SKU"].astype(str).apply(_normalizar_sku_exportacao).tolist())
        selected_skus, selection_table = _render_editor_exportacao(result.analise, f"comum_editor_{payload['generation_id']}", default_skus)
        filtered_result = _build_common_filtered_result(result, selected_skus)

        c1, c2 = st.columns(2)
        c1.metric("SKUs selecionados", len(selected_skus))
        c2.metric("Tabelas Protheus selecionadas", len(payload.get("selecionados", {}).get("tabelas", [])))

        outputs, zip_name, zip_bytes = _build_common_outputs(payload, filtered_result)

        st.subheader("4. Baixar arquivos selecionados")
        _render_downloads(outputs, zip_name, zip_bytes, "comum")

        tabs = st.tabs(["Análise", "Log", "001", "004", "012", "007", "013"])
        with tabs[0]:
            st.dataframe(filtered_result.analise, use_container_width=True, hide_index=True)
        with tabs[1]:
            st.dataframe(result.log, use_container_width=True, hide_index=True)
            st.text_area("Prévia do log TXT", result.log_text, height=280, key="comum_log_preview")
        tabelas_selecionadas_payload = set(payload.get("selecionados", {}).get("tabelas", []))
        for idx, tabela in enumerate(["001", "004", "012", "007", "013"], start=2):
            with tabs[idx]:
                if tabela in tabelas_selecionadas_payload:
                    st.dataframe(filtered_result.saidas.get(tabela), use_container_width=True, hide_index=True)
                else:
                    st.write(f"Tabela {tabela} não foi selecionada para exportação nesta execução.")


if modulo == "Precificação Comum":
    render_precificacao_comum()
else:
    render_adequacao_ml()
