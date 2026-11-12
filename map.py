from pathlib import Path
import pandas as pd
import re
from decimal import Decimal, InvalidOperation

caminho = Path(__file__).parent / "data"
output = Path(__file__).parent / "output"
vtex = caminho / "estoque.xls"
gix = caminho / "gix.csv"


# Arquivos
gix = pd.read_csv(
    gix, 
    sep=";",
    dtype=str, 
    encoding="cp1252"
)
estoque = pd.concat(
    pd.read_excel(
        vtex,
        sheet_name=None,
        dtype=str
    ).values(),
    ignore_index=True
)

# Normalização EAN-13
def normalizar_ean(valor):
    if pd.isna(valor):
        return None

    try:
        numero = Decimal(str(valor).strip().replace(",", "."))

        if numero != numero.to_integral_value():
            return None

        ean = str(int(numero))

        if not re.fullmatch(r"\d{13}", ean):
            return None

        return ean

    except (InvalidOperation, ValueError):
        return None


# Preparar chaves
gix["EAN"] = gix["CÓDIGOBARRAS"].apply(normalizar_ean)
estoque["EAN"] = estoque["RefId"].apply(normalizar_ean)

# Descartar EANs inválidos
gix = gix.dropna(subset=["EAN"])
estoque = estoque.dropna(subset=["EAN"])

# Verificar duplicidades na VTEX
if estoque["EAN"].duplicated().any():
    raise ValueError("EAN duplicado na base VTEX!")

# Cruzar dados
resultado = pd.merge(
    gix,
    estoque,
    on="EAN",
    how="inner",
    validate="many_to_one"
)

# Gerar mapa
colunas_mapa = [
    "CÓDIGO",
    "SkuId",
    "EAN",
    "DESCRIÇÃO"
]

mapa = resultado[colunas_mapa].copy()

mapa.columns = [
    "gix",
    "SkuId",
    "EAN-13",
    "Descrição"
]

output.mkdir(parents=True, exist_ok=True)

# Exportar CSV
mapa.to_csv(
    output / "mapa.csv",
    index=False,
    sep=";",
    encoding="utf-8-sig"
)

print(f"Mapa gerado: {len(mapa)} correspondências")