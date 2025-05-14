import os
from barcode import get as get_barcode
from barcode.writer import ImageWriter

def generate_gtin_barcodes(df, output_folder="barcodes"):
    os.makedirs(output_folder, exist_ok=True)
    df["StrekkodeFil"] = ""

    for idx, row in df.iterrows():
        gtin = row.get("GTIN", "")
        if gtin and gtin.isdigit() and len(gtin) == 13:
            try:
                path = os.path.join(output_folder, f"gtin_{idx}")
                ean = get_barcode('ean13', gtin, writer=ImageWriter())
                img_path = ean.save(path)
                df.at[idx, "StrekkodeFil"] = img_path
            except Exception as ex:
                print(f"Strekkodefeil for {gtin}: {ex}")
    return df
