from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from processor import (
    ProcessConfig,
    processar_precos,
    build_report_excel,
    build_protheus_csv,
    build_zip,
    build_log_txt,
    CommonPricingConfig,
    processar_precificacao_comum,
    build_common_report_excel,
    build_common_protheus_csv,
    build_common_log_txt,
    build_zip_generic,
)


def run_ml(args):
    out = Path(args.saida)
    out.mkdir(parents=True, exist_ok=True)

    cfg = ProcessConfig(
        tabela_codigo=str(args.codigo).zfill(3),
        tabela_codigo_013=str(args.codigo_013).zfill(3),
        acrescimo_tabela_013=args.acrescimo_013,
        desconto_ml=args.desconto,
    )
    result = processar_precos(args.tabela_007, args.ml, args.anterior, cfg)

    data_nome = datetime.now().strftime("%Y%m%d_%H%M")
    report_name = f"relatorio_tabela_007_013_ml_{data_nome}.xlsx"
    protheus_007_name = f"protheus_tabela_{cfg.tabela_codigo}_{data_nome}.csv"
    protheus_013_name = f"protheus_tabela_{cfg.tabela_codigo_013}_{data_nome}.csv"
    log_name = f"log_tabela_007_013_ml_{data_nome}.txt"
    zip_name = f"pacote_tabela_007_013_ml_{data_nome}.zip"

    report_bytes = build_report_excel(result)
    protheus_007_bytes = build_protheus_csv(result, cfg, tabela=cfg.tabela_codigo)
    protheus_013_bytes = build_protheus_csv(result, cfg, tabela=cfg.tabela_codigo_013)
    log_txt_bytes = build_log_txt(result)
    zip_bytes = build_zip(report_bytes, protheus_007_bytes, protheus_013_bytes, log_txt_bytes, report_name, protheus_007_name, protheus_013_name, log_name)

    files = {
        report_name: report_bytes,
        protheus_007_name: protheus_007_bytes,
        protheus_013_name: protheus_013_bytes,
        log_name: log_txt_bytes,
        zip_name: zip_bytes,
    }
    for name, content in files.items():
        (out / name).write_bytes(content)

    print("Arquivos ML gerados em:", out)


def run_comum(args):
    out = Path(args.saida)
    out.mkdir(parents=True, exist_ok=True)

    cfg = CommonPricingConfig(
        desconto_007_sobre_001=args.desconto_007,
        acrescimo_013_sobre_007=args.acrescimo_013,
    )
    result = processar_precificacao_comum(args.base, config=cfg)

    data_nome = datetime.now().strftime("%Y%m%d_%H%M")
    report_name = f"relatorio_precificacao_comum_{data_nome}.xlsx"
    log_name = f"log_precificacao_comum_{data_nome}.txt"
    zip_name = f"pacote_precificacao_comum_{data_nome}.zip"

    files = {
        report_name: build_common_report_excel(result),
        log_name: build_common_log_txt(result),
    }
    for tabela in ["001", "004", "012", "007", "013"]:
        files[f"protheus_tabela_{tabela}_{data_nome}.csv"] = build_common_protheus_csv(result, tabela)

    files[zip_name] = build_zip_generic(files)

    for name, content in files.items():
        (out / name).write_bytes(content)

    print("Arquivos de Precificação Comum gerados em:", out)


def main():
    parser = argparse.ArgumentParser(description="GERADOR DE TABELAS DE PREÇOS EM MASSA.")
    sub = parser.add_subparsers(dest="modo", required=True)

    ml = sub.add_parser("ml", help="Adequação ML para tabelas 007/013")
    ml.add_argument("--tabela-007", required=True, help="Caminho da tabela 007 oficial .xlsx")
    ml.add_argument("--ml", required=True, help="Caminho da planilha PREÇOS ML - REVISADO .xlsx")
    ml.add_argument("--anterior", default=None, help="Relatório anterior gerado pelo app, opcional")
    ml.add_argument("--saida", default="saida_tabela_007_ml", help="Pasta de saída")
    ml.add_argument("--codigo", default="007", help="Código da tabela Protheus base")
    ml.add_argument("--codigo-013", default="013", help="Código da tabela Protheus com acréscimo")
    ml.add_argument("--acrescimo-013", type=float, default=0.12, help="Acréscimo da tabela 013 sobre a 007. Ex.: 0.12")
    ml.add_argument("--desconto", type=float, default=0.10, help="Desconto sobre preço ML. Ex.: 0.10")
    ml.set_defaults(func=run_ml)

    comum = sub.add_parser("comum", help="Precificação comum para tabelas 001/004/012/007/013")
    comum.add_argument("--base", required=True, help="Planilha base com A=SKU, B=Custo, C=Preço 001")
    comum.add_argument("--saida", default="saida_precificacao_comum", help="Pasta de saída")
    comum.add_argument("--desconto-007", type=float, default=0.16, help="Desconto da 007 sobre a 001. Ex.: 0.16")
    comum.add_argument("--acrescimo-013", type=float, default=0.12, help="Acréscimo da 013 sobre a 007. Ex.: 0.12")
    comum.set_defaults(func=run_comum)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
