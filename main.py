import os.path

import pandas as pd
from glob import glob
import os
import shutil

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)


def verify_file(xlsfile: str, output_base:str):
    print(f"Verifica in corso per il file {xlsfile}")
    # xlsfile ='/Users/bruand/Desktop/AppVotoZone/test_app_voto_zona_hirpinia.xls'
    outfile = os.path.join(output_base,f"{os.path.splitext(os.path.basename(xlsfile))[0]}.xlsx")
    df = pd.read_excel(xlsfile)
    try:
        df["double"] = df["EMAIL"].str.contains(";")
        duplicated = df["EMAIL"].duplicated()
        duplicated_mail = df["EMAIL"].loc[duplicated].tolist()
        df["duplicated"] = df["EMAIL"].isin(duplicated_mail)

        df["check"] = ~(df["double"] | df["duplicated"])
        df["double"] = ~df["double"]
        df["duplicated"] = ~df["duplicated"]
        df.to_excel(outfile, index=False)
    except KeyError:
        print("Manca colonna EMAIL, file non valido")
    print(f"Verifica completata")


if __name__ == '__main__':
    input_path = "input"
    output_dir = "output"
    inpathsxls = os.path.join(input_path, "*.[xX][lL][sS]")
    inpathsxlsx = os.path.join(input_path, "*.[xX][lL][sS][xX]")
    xls_files = [f for f in glob(inpathsxls)] + [f for f in glob(inpathsxlsx)]
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    [verify_file(f, output_dir) for f in xls_files]
