import pandas as pd
import os
from common.paths import DATA_PATH

def analyze_deadline_with_reason(sheet_name="Sheet1", file_name="Dados Compesa (Comercial).xlsx"):
    file_path = os.path.join(DATA_PATH, file_name)
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    df["Dt Solicitação RA"] = pd.to_datetime(df["Dt Solicitação RA"], errors="coerce")
    df["Dt Encerramento RA"] = pd.to_datetime(df["Dt Encerramento RA"], errors="coerce")
    df["Prazo Tipo Sol RA"] = pd.to_numeric(df["Prazo Tipo Sol RA"], errors="coerce")

    valid_df = df.dropna(subset=["Dt Solicitação RA", "Dt Encerramento RA", "Prazo Tipo Sol RA"]).copy()

    valid_df["Dias Decorridos"] = (valid_df["Dt Encerramento RA"] - valid_df["Dt Solicitação RA"]).dt.days

    valid_df["Situação Prazo"] = valid_df.apply(
        lambda row: "Dentro do Prazo" if row["Dias Decorridos"] <= row["Prazo Tipo Sol RA"] else "Fora do Prazo", axis=1
    )


    total_valid_rows = len(valid_df)
    within_deadline = (valid_df["Situação Prazo"] == "Dentro do Prazo").sum()
    outside_deadline = (valid_df["Situação Prazo"] == "Fora do Prazo").sum()

    pct_within = (within_deadline / total_valid_rows * 100) if total_valid_rows else 0
    pct_outside = (outside_deadline / total_valid_rows * 100) if total_valid_rows else 0

    reasons_counts = valid_df[valid_df["Situação Prazo"] == "Fora do Prazo"]["Motivo Encer RA"].value_counts(dropna=False).to_dict()

    return {
        "Quantidade total de atendimentos": total_valid_rows,
        "Quantidade dentro do prazo": within_deadline,
        "Quantidade fora do prazo": outside_deadline,
        "% dentro do prazo": round(pct_within, 2),
        "% fora do prazo": round(pct_outside, 2),
        "Contagem Motivos Fora do Prazo": reasons_counts
    }