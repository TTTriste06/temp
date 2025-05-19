import os
import pandas as pd
from io import BytesIO
from pivot_processor import PivotProcessor
from config import CONFIG

def log(msg):
    print(f"[INFO] {msg}")

def main():
    # 你可以在这里覆盖 config.py 中的路径（推荐改这里用于测试）
    input_dir = CONFIG["input_dir"]  # 或手动写："/Users/tttriste.kkkkkk/Desktop/semi/原始数据"
    output_file = CONFIG["output_file"]  # 自动生成时间戳命名的输出文件路径

    log(f"📂 读取目录: {input_dir}")
    log(f"📄 目标输出文件: {output_file}")

    core_filenames = [
        "赛卓-未交订单.xlsx",
        "赛卓-成品在制.xlsx",
        "赛卓-成品库存.xlsx",
        "赛卓-晶圆库存.xlsx",
        "赛卓-CP在制.xlsx"
    ]

    auxiliary_filenames = {
        "赛卓-预测.xlsx": "forecast_file",
        "赛卓-安全库存.xlsx": "safety_file",
        "赛卓-新旧料号.xlsx": "mapping_file"
    }

    uploaded_files = {}
    additional_sheets = {}

    # 读取主数据文件
    for filename in core_filenames:
        file_path = os.path.join(input_dir, filename)
        if os.path.exists(file_path):
            log(f"✅ 加载主数据文件: {filename}")
            uploaded_files[filename] = open(file_path, "rb")
        else:
            log(f"❌ 未找到主数据文件: {filename}")

    if len(uploaded_files) < 5:
        log("❌ 主数据不足 5 个，退出。")
        return

    # 读取辅助数据
    for filename in auxiliary_filenames:
        file_path = os.path.join(input_dir, filename)
        if os.path.exists(file_path):
            log(f"✅ 加载辅助数据: {filename}")
            df = pd.read_excel(file_path, header=None if filename == "赛卓-新旧料号.xlsx" else 0)
            additional_sheets[filename.replace(".xlsx", "")] = df
        else:
            log(f"⚠️ 未找到辅助数据文件: {filename}")

    # 执行数据处理
    output_buffer = BytesIO()
    processor = PivotProcessor()
    processor.process(uploaded_files, output_buffer, additional_sheets)

    # 保存为本地 Excel 文件
    with open(output_file, "wb") as f:
        f.write(output_buffer.getvalue())
        log(f"📤 汇总成功！已保存至: {output_file}")

if __name__ == "__main__":
    main()
