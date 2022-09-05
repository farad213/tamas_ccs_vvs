import pandas as pd

"""Excel file has to be named 'BMW.xlsx'
if there is a new sheet, make sure to update 'sheet_name' in code (line 6)"""

df = pd.read_excel("BMW.xlsx", sheet_name="2022.07.29.", header=1, usecols=[4, 5, 7, 8],
                   names=["ccs notation", "ccs length", "vvs notation", "vvs length"])


def get_length(to_check, type=None):
    assert type in ["ccs", "vvs"]
    assert isinstance(to_check, str)

    index = df[f"{type} notation"][df[f"{type} notation"] == to_check].index
    result = df[f"{type} length"][index].values
    if list(result):
        return result[0]
    return f"Notation not found in {type} column"


def main():
    ccs = (input("Cölöpcsúcs: ")).strip()
    ccs_length = get_length(ccs, type="ccs")
    if isinstance(ccs_length, str):
        print("Érvénytelen cölöpcsúcs jelölés")
        return main()

    vvs = (input("Visszavésési szint: ")).strip()
    vvs_length = get_length(vvs, type="vvs")
    if isinstance(vvs_length, str):
        print("Érvénytelen visszavésési szint jelölés")
        return main()

    result = round(vvs_length - ccs_length, 2)
    print(result)


if __name__ == "__main__":
    while True:
        main()
