import threading
import pandas as pd
from utils.barcode_utils import generate_gtin_barcodes as real_gen

def generate_gtin_with_progress(df: pd.DataFrame, progress_func=None):
    def _update_progress(value, pct):
        if progress_func:
            progress_func(pct)

    result = {"data": None}

    def _worker():
        dfs = []
        total = len(df)
        for idx, row in enumerate(df.itertuples(index=False), start=1):
            one = pd.DataFrame([row._asdict()])
            one = real_gen(one)
            dfs.append(one)

            pct = int(idx / total * 100)
            _update_progress(idx, pct)

        result["data"] = pd.concat(dfs, ignore_index=True)

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
    thread.join()

    return result["data"]
