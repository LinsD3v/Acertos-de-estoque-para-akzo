from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata

import pandas as pd
import xlwt

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
OUTPUT_FILE = OUTPUT_DIR / "estoque_ajuste.csv"
OUTPUT_XLS = OUTPUT_DIR / "estoque_ajuste.xls"

CAMPOS = [
    "SkuId",
    "TotalQuantity",
    "ReservedQuantity",
    "AvailableQuantity",
    "WarehouseId",
    "WarehouseName",
    "RefId",
    "IsActive",
    "UnlimitedQuantity",
    "LockId",
    "DispatchedReservations",
]


def normalizar_coluna(nome: object) -> str:
    texto = unicodedata.normalize("NFKD", str(nome))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_ean(valor: object) -> str | None:
    if pd.isna(valor):
        return None

    texto = str(valor).strip()
    if not texto:
        return None

    numeros = "".join(ch for ch in texto if ch.isdigit())
    if len(numeros) != 13:
        return None

    return numeros


def parse_quantidade(valor: object) -> int:
    if pd.isna(valor):
        return 0

    texto = str(valor).strip()
    if not texto:
        return 0

    texto = texto.replace(".", "").replace(",", ".")

    try:
        return int(Decimal(texto))
    except (InvalidOperation, ValueError):
        return 0


def gerar_mapa_se_necessario() -> Path:
    mapa_path = OUTPUT_DIR / "mapa.csv"
    if mapa_path.exists():
        try:
            mapa_preview = pd.read_csv(mapa_path, sep=";", encoding="utf-8-sig", dtype=str, nrows=1)
            mapa_preview.columns = [normalizar_coluna(col) for col in mapa_preview.columns]
            if {"SkuId", "EAN-13"}.issubset(mapa_preview.columns):
                return mapa_path
        except Exception:
            pass

    gix_path = DATA_DIR / "gix.csv"
    estoque_path = DATA_DIR / "estoque.xls"

    if not gix_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {gix_path}")
    if not estoque_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {estoque_path}")

    gix = pd.read_csv(gix_path, sep=";", encoding="cp1252", dtype=str)
    gix.columns = [normalizar_coluna(col) for col in gix.columns]
    gix["EAN"] = gix["CODIGOBARRAS"].apply(normalizar_ean)
    gix = gix.dropna(subset=["EAN"]).copy()

    estoque = pd.concat(
        pd.read_excel(estoque_path, sheet_name=None, dtype=str).values(),
        ignore_index=True,
    )
    estoque.columns = [normalizar_coluna(col) for col in estoque.columns]
    estoque["EAN"] = estoque["RefId"].apply(normalizar_ean)
    estoque = estoque.dropna(subset=["EAN"]).copy()

    if estoque["EAN"].duplicated().any():
        raise ValueError("EAN duplicado na base VTEX!")

    resultado = pd.merge(
        gix,
        estoque,
        on="EAN",
        how="inner",
        validate="many_to_one",
    )

    mapa = resultado[["CODIGO", "SkuId", "EAN", "DESCRICAO"]].copy()
    mapa.columns = ["gix", "SkuId", "EAN-13", "Descrição"]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    mapa.to_csv(mapa_path, index=False, sep=";", encoding="utf-8-sig")
    return mapa_path


def gerar_estoque_importacao() -> pd.DataFrame:
    mapa_path = gerar_mapa_se_necessario()
    mapa = pd.read_csv(mapa_path, sep=";", encoding="utf-8-sig", dtype=str)
    mapa.columns = [normalizar_coluna(col) for col in mapa.columns]
    if "EAN-13" not in mapa.columns:
        if "EAN 13" in mapa.columns:
            mapa = mapa.rename(columns={"EAN 13": "EAN-13"})
        else:
            raise KeyError("Coluna EAN-13 não encontrada no mapa de relacionamento.")
    mapa = mapa[["SkuId", "EAN-13"]].copy()
    mapa["SkuId"] = mapa["SkuId"].astype(str).str.strip()
    mapa["EAN"] = mapa["EAN-13"].astype(str).str.strip()
    mapa = mapa[["SkuId", "EAN"]].drop_duplicates(subset=["EAN"]).copy()

    gix = pd.read_csv(DATA_DIR / "gix.csv", sep=";", encoding="cp1252", dtype=str)
    gix.columns = [normalizar_coluna(col) for col in gix.columns]
    gix["EAN"] = gix["CODIGOBARRAS"].apply(normalizar_ean)
    gix = gix.dropna(subset=["EAN"]).copy()
    gix["TotalQuantity"] = gix["ESTOQUEATUAL"].apply(parse_quantidade)

    resultado = gix[["EAN", "TotalQuantity"]].merge(
        mapa,
        on="EAN",
        how="inner",
        validate="many_to_one",
    )

    final = pd.DataFrame({
        "SkuId": resultado["SkuId"].astype(str),
        "TotalQuantity": resultado["TotalQuantity"].astype(int),
        "ReservedQuantity": 0,
        "AvailableQuantity": resultado["TotalQuantity"].astype(int),
        "WarehouseId": "1_1",
        "WarehouseName": "Estoque principal",
        "RefId": resultado["EAN"].astype(str),
        "IsActive": True,
        "UnlimitedQuantity": False,
        "LockId": None,
        "DispatchedReservations": 0,
    })

    final = final[CAMPOS]
    return final


def exportar_xls(df: pd.DataFrame, caminho: Path) -> None:
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("estoque")

    for coluna_idx, nome_coluna in enumerate(df.columns):
        sheet.write(0, coluna_idx, str(nome_coluna))

    for linha_idx, linha in enumerate(df.itertuples(index=False, name=None), start=1):
        for coluna_idx, valor in enumerate(linha):
            if pd.isna(valor):
                valor = ""
            sheet.write(linha_idx, coluna_idx, valor)

    workbook.save(str(caminho))


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    final = gerar_estoque_importacao()
    final.to_csv(OUTPUT_FILE, index=False, sep=";", encoding="utf-8-sig", na_rep="")
    exportar_xls(final, OUTPUT_XLS)

    print(f"Arquivo CSV gerado: {OUTPUT_FILE}")
    print(f"Arquivo XLS gerado: {OUTPUT_XLS}")
    print(f"Linhas exportadas: {len(final)}")


if __name__ == "__main__":
    main()
