from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
MASTER_FILE = "master.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file:
        return "没有上传文件"

    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_input.xlsx")
    output_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_output.xlsx")

    file.save(input_path)

    # 读取主数据
    master_df = pd.read_excel(MASTER_FILE)

    # 清洗主数据
    master_df["物料编号"] = master_df["物料编号"].astype(str).str.strip()

    # 构建映射：物料编号 -> 净重
    weight_map = master_df.set_index("物料编号")["净重"]

    # 读取所有 sheet
    all_sheets = pd.read_excel(input_path, sheet_name=None)

    result_sheets = {}

    for sheet_name, df in all_sheets.items():
        # 如果没有“物料编号”列，直接原样返回
        if "物料编号" not in df.columns:
            result_sheets[sheet_name] = df
            continue

        df["物料编号"] = df["物料编号"].astype(str).str.strip()

        # 如果没有“净重”列，先创建
        if "净重" not in df.columns:
            df["净重"] = None

        # 核心逻辑：查表 + 保留原值
        new_weight = df["物料编号"].map(weight_map)

        df["净重"] = new_weight.fillna(df["净重"])

        result_sheets[sheet_name] = df

    # 写入多 sheet Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in result_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)