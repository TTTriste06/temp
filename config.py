from datetime import datetime

CONFIG = {
    "input_dir": r"/Users/tttriste.kkkkkk/Desktop/semi",
    "output_file": f"/Users/tttriste.kkkkkk/Desktop/semi/运营数据订单-在制-库存汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    "pivot_config": {
        "weijiaodindan.xlsx": {
            "index": ["晶圆品名", "规格", "品名"],
            "columns": "预交货日",
            "values": ["订单数量", "未交订单数量"],
            "aggfunc": "sum",
            "date_format": "%Y-%m"
        },
        "chengpinzaizhi.xlsx": {
            "index": ["工作中心", "封装形式", "晶圆型号", "产品规格", "产品品名"],
            "columns": "预计完工日期",
            "values": ["未交"],
            "aggfunc": "sum",
            "date_format": "%Y-%m"
        },
        "CPzaizhi.xlsx": {
            "index": ["晶圆型号", "产品品名"],
            "columns": "预计完工日期",
            "values": ["未交"],
            "aggfunc": "sum",
            "date_format": "%Y-%m"
        },
        "chengpinkucun.xlsx": {
            "index": ["WAFER品名", "规格", "品名"],
            "columns": "仓库名称",
            "values": ["数量"],
            "aggfunc": "sum"
        },
        "jingyuankucun.xlsx": {
            "index": ["WAFER品名", "规格"],
            "columns": "仓库名称",
            "values": ["数量"],
            "aggfunc": "sum"
        }
    }
}

# config.py 中新增
FILE_RENAME_MAPPING = {
    "赛卓-未交订单.xlsx": "unfulfilled_orders.xlsx",
    "赛卓-成品在制.xlsx": "finished_products.xlsx",
    "赛卓-成品库存.xlsx": "finished_inventory.xlsx",
    "赛卓-晶圆库存.xlsx": "wafer_inventory.xlsx",
    "赛卓-CP在制.xlsx": "cp_in_progress.xlsx",
    "赛卓-预测.xlsx": "forecast.xlsx",
    "赛卓-安全库存.xlsx": "safety.xlsx",
    "赛卓-新旧料号.xlsx": "mapping.xlsx"
}

# 用于生成 sheet 名时还原中文
REVERSE_MAPPING = {v.replace(".xlsx", ""): k.replace(".xlsx", "") for k, v in FILE_RENAME_MAPPING.items()}

