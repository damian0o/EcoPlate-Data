import pandas as pd
import json
import glob
import os

def convert_xlsx_to_json(xlsx_path):
    wave_590 = pd.read_excel(xlsx_path, usecols='B:M', skiprows=lambda x: x < 5 or x > 13, header=None)
    wave_720 = pd.read_excel(xlsx_path, usecols='B:M', skiprows=lambda x: x < 18 or x > 26, header=None)
    wave = round((wave_590 - wave_720), 3)
    matrices = []
    for start_col in range(0, 12, 4):
        matrix = wave.iloc[:, start_col:start_col + 4].values.tolist()
        matrices.append(matrix)
    return {"filename": os.path.basename(xlsx_path), "matrices": matrices}

def main():
    xlsx_files = glob.glob("*.xlsx")
    index = []
    for xlsx_path in xlsx_files:
        basename = os.path.splitext(os.path.basename(xlsx_path))[0]
        json_filename = f"{basename}.json"
        data = convert_xlsx_to_json(xlsx_path)
        with open(json_filename, "w") as f:
            json.dump(data, f, indent=2)
        index.append({"name": json_filename, "source": os.path.basename(xlsx_path)})
    with open("index.json", "w") as f:
        json.dump(index, f, indent=2)

if __name__ == "__main__":
    main()
