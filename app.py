from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import uuid
import shutil

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
MASTER_FILE = "master.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def clear_upload_folder():
    """清空 uploads 目录"""
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"删除失败: {file_path}, {e}")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file:
        return "没有上传文件"

    # 🧹 先清空 uploads 目录
    clear_upload_folder()

    # 👉 原始文件名
    original_filename = file.filename
    name, ext = os.path.splitext(original_filename)

    file_id = str(uuid.uuid4())

    input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_input{ext}")
    output_filename = f"{name}{ext}"
    output_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{output_filename}")

    file.save(input_path)

    # ===== 主数据处理 =====
    master_df = pd.read_excel(MASTER_FILE)
    master_df["物料编号"] = master_df["物料编号"].astype(str).str.strip()

    weight_map = master_df.set_index("物料编号")["净重"]

    # ===== 多 sheet 处理 =====
    all_sheets = pd.read_excel(input_path, sheet_name=None)
    result_sheets = {}

    for sheet_name, df in all_sheets.items():
        if "物料编号" not in df.columns:
            result_sheets[sheet_name] = df
            continue

        df["物料编号"] = df["物料编号"].astype(str).str.strip()

        if "净重" not in df.columns:
            df["净重"] = None

        # 查表并保留原值
        new_weight = df["物料编号"].map(weight_map)
        df["净重"] = new_weight.fillna(df["净重"])

        result_sheets[sheet_name] = df

    # ===== 写入输出文件（多 sheet）=====
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in result_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # ===== 返回下载 =====
    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_filename
    )


if __name__ == "__main__":
    app.run(host="localhost", port=80)