import streamlit as st
import pandas as pd
import json
import os
import sys
import io
import urllib.request
import urllib.error
from datetime import datetime

# ==================== 页面全局配置 ====================
st.set_page_config(
    page_title="货件 PCS 计算工具",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================== 配置管理 ====================
def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.path.abspath('.')

CONFIG_FILE_PATH = os.path.join(get_app_dir(), "columns_config.json")

DEFAULT_CONFIG = {
    "airscript": {
        "webhook_url": "https://www.kdocs.cn/api/v3/ide/file/cdH0A450EedY/script/V2-6x7HgWruLz4P74YP1PIsVf/sync_task",
        "token": "RRZwCZarLOHD4gHKtfSVi"
    }
}

def load_columns_config():
    if not os.path.exists(CONFIG_FILE_PATH):
        try:
            with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return DEFAULT_CONFIG

    try:
        with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            if "airscript" not in cfg:
                cfg["airscript"] = DEFAULT_CONFIG["airscript"]
            return cfg
    except Exception:
        return DEFAULT_CONFIG

# ==================== AirScript 数据拉取 ====================
def parse_airscript_response(res_data):
    if isinstance(res_data, dict):
        if "data" in res_data and isinstance(res_data["data"], dict) and "result" in res_data["data"]:
            result = res_data["data"]["result"]
        elif "result" in res_data:
            result = res_data["result"]
        else:
            result = res_data
    else:
        result = res_data

    if isinstance(result, list):
        if len(result) == 0:
            return pd.DataFrame()
        first_item = result[0]
        if isinstance(first_item, dict):
            if 'fields' in first_item and isinstance(first_item['fields'], dict):
                records = [item.get('fields', {}) for item in result if isinstance(item, dict)]
                return pd.DataFrame(records)
            return pd.DataFrame(result)
        elif isinstance(first_item, (list, tuple)):
            headers = [str(h).strip() for h in first_item]
            rows = result[1:]
            return pd.DataFrame(rows, columns=headers)
    elif isinstance(result, dict):
        if 'records' in result and isinstance(result['records'], list):
            return parse_airscript_response(result['records'])
        return pd.DataFrame([result])

    return pd.DataFrame()

def fetch_purchase_from_airscript(url, token, timeout=35):
    headers = {
        'AirScript-Token': token.strip(),
        'Content-Type': 'application/json; charset=utf-8',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
    }
    payload = json.dumps({"Context": {"argv": {}}}).encode('utf-8')
    req = urllib.request.Request(url.strip(), data=payload, headers=headers, method='POST')

    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            status_code = response.getcode()
            resp_body = response.read().decode('utf-8')
            if status_code != 200:
                raise Exception(f"HTTP 请求失败，状态码: {status_code}")
            
            json_data = json.loads(resp_body)
            if isinstance(json_data, dict) and json_data.get("status") == "error":
                msg = json_data.get("message") or json_data.get("msg") or str(json_data)
                raise Exception(f"AirScript 执行报错: {msg}")
            
            df = parse_airscript_response(json_data)
            if df.empty:
                raise Exception("AirScript 响应成功，但返回数据为空！")
            return df
    except urllib.error.HTTPError as he:
        raise Exception(f"网络请求失败 [HTTP {he.code}]: {he.reason}")
    except urllib.error.URLError as ue:
        raise Exception(f"网络连接超时或无法连接: {ue.reason}")
    except Exception as e:
        raise e

# ==================== Excel 导出 ====================
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ==================== 主界面 ====================
def main():
    st.title("📦 货件 PCS 计算工具")
    st.caption("上传发货单 Excel")
    st.divider()

    # 1. 上传文件（上传后自动执行）
    uploaded_file = st.file_uploader("请选择发货单 Excel 文件（.xlsx）", type=["xlsx"])

    if uploaded_file is not None:
        with st.spinner("正在拉取云端采购单并计算数据，请稍候..."):
            try:
                cfg = load_columns_config()
                air_cfg = cfg.get("airscript", DEFAULT_CONFIG["airscript"])
                webhook_url = air_cfg.get("webhook_url", DEFAULT_CONFIG["airscript"]["webhook_url"])
                token = air_cfg.get("token", DEFAULT_CONFIG["airscript"]["token"])

                # 拉取云端采购单
                purchase_df = fetch_purchase_from_airscript(webhook_url, token)
                purchase_df.columns = [str(c).strip() for c in purchase_df.columns]

                # 兼容品名与PCS列名
                if "中文品名" not in purchase_df.columns and "品名" in purchase_df.columns:
                    purchase_df["中文品名"] = purchase_df["品名"]
                
                if "PCS" not in purchase_df.columns:
                    if "单箱数量" in purchase_df.columns:
                        purchase_df["PCS"] = purchase_df["单箱数量"]
                    elif "箱规" in purchase_df.columns:
                        purchase_df["PCS"] = purchase_df["箱规"]

                # 读取发货单
                delivery_df = pd.read_excel(uploaded_file)
                delivery_df.columns = [str(c).strip() for c in delivery_df.columns]

                required_purchase_cols = ["SKU", "中文品名", "PCS"]
                required_delivery_cols = ["SKU", "发货量", "货件编号"]

                for col in required_purchase_cols:
                    if col not in purchase_df.columns:
                        raise ValueError(f"云端采购单缺少必需列：【{col}】")

                for col in required_delivery_cols:
                    if col not in delivery_df.columns:
                        raise ValueError(f"发货单缺少必需列：【{col}】")

                # SKU 匹配与计算
                purchase_df["SKU"] = purchase_df["SKU"].astype(str).str.strip()
                delivery_df["SKU"] = delivery_df["SKU"].astype(str).str.strip()

                purchase_mapping = purchase_df[["SKU", "中文品名", "PCS"]].drop_duplicates(subset=["SKU"])

                delivery_updated = pd.merge(
                    delivery_df,
                    purchase_mapping,
                    on="SKU",
                    how="left"
                )

                delivery_updated["PCS"] = pd.to_numeric(delivery_updated["PCS"], errors="coerce").fillna(0)
                delivery_updated["发货量"] = pd.to_numeric(delivery_updated["发货量"], errors="coerce").fillna(0)
                delivery_updated["总PCS"] = delivery_updated["发货量"] * delivery_updated["PCS"]

                # 生成货件汇总表
                summary_data = delivery_updated.groupby(["货件编号", "中文品名"]).agg({
                    "总PCS": "sum"
                }).reset_index()[["货件编号", "中文品名", "总PCS"]]

                # 检查是否存在 PCS 为 0 的异常项
                zero_pcs_df = delivery_updated[delivery_updated["PCS"] == 0]
                if len(zero_pcs_df) > 0:
                    st.error(f"⚠️ **检测到共有 {len(zero_pcs_df)} 条记录的 PCS 为 0（未匹配到采购单或单箱数量为0）：**")
                    st.dataframe(zero_pcs_df[["SKU", "货件编号", "发货量", "PCS"]], use_container_width=True)

                # 下载按钮（仅保留货件汇总表）
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_summary_bytes = to_excel_bytes(summary_data)

                st.download_button(
                    label="📥 下载货件汇总表 (.xlsx)",
                    data=file_summary_bytes,
                    file_name=f"货件编号_总PCS汇总表_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

                # 直接平铺展示货件汇总表
                st.subheader("📊 货件汇总表")
                st.dataframe(summary_data, use_container_width=True)

            except Exception as e:
                st.error(f"❌ 处理失败: {str(e)}")

if __name__ == "__main__":
    main()