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
    page_title="货件 PCS 计算与汇总工具",
    page_icon="📦",
    layout="wide"
)

# ==================== 路径与配置管理 ====================
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
    """读取或初始化配置文件"""
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

def save_columns_config(cfg):
    """保存配置到本地 json 文件"""
    try:
        with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.sidebar.error(f"保存配置失败: {e}")
        return False

# ==================== AirScript 数据拉取 ====================
def parse_airscript_response(res_data):
    """解析并标准化 AirScript 返回的多种数据结构为 DataFrame"""
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
    """通过 Webhook POST 请求 AirScript 获取采购单数据"""
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

# ==================== Excel 导出工具 ====================
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def to_combined_excel_bytes(df_delivery: pd.DataFrame, df_summary: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_delivery.to_excel(writer, sheet_name='已更新发货单', index=False)
        df_summary.to_excel(writer, sheet_name='货件汇总表', index=False)
    return output.getvalue()

# ==================== 主页面交互 ====================
def main():
    # 标题部分
    st.title("📦 货件 PCS 计算工具 (在线 Web 版)")
    st.markdown("上传发货单 Excel 文件，自动拉取金山文档 AirScript 云端采购单数据，匹配品名、单箱数量并计算汇总总 PCS。")
    st.divider()

    # 加载配置
    cfg = load_columns_config()
    air_cfg = cfg.get("airscript", DEFAULT_CONFIG["airscript"])

    # 侧边栏：接口与参数配置
    with st.sidebar:
        st.header("⚙️ 云端接口配置")
        webhook_url_input = st.text_input(
            "AirScript Webhook URL", 
            value=air_cfg.get("webhook_url", DEFAULT_CONFIG["airscript"]["webhook_url"])
        )
        token_input = st.text_input(
            "AirScript Token", 
            value=air_cfg.get("token", DEFAULT_CONFIG["airscript"]["token"]), 
            type="password"
        )
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("💾 保存配置", use_container_width=True):
                cfg["airscript"] = {"webhook_url": webhook_url_input, "token": token_input}
                if save_columns_config(cfg):
                    st.success("配置已更新！")
        
        with col_btn2:
            test_clicked = st.button("🔄 测试连接", use_container_width=True)

        if test_clicked:
            with st.spinner("正在连接云端..."):
                try:
                    test_df = fetch_purchase_from_airscript(webhook_url_input, token_input)
                    st.success(f"✅ 连接成功！共拉取到 {len(test_df)} 条采购单记录。")
                except Exception as ex:
                    st.error(f"❌ 连接失败：{str(ex)}")

    # 主体布局：分为两步
    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("1. 云端采购单状态")
        st.info("数据源已配置为金山文档 AirScript Webhook 接口。")
        pull_online = st.checkbox("处理时实时同步最新采购单数据", value=True)

    with col_right:
        st.subheader("2. 上传发货单文件")
        uploaded_file = st.file_uploader("请选择发货单 Excel 文件（.xlsx）", type=["xlsx"])

    # 如果上传了文件，提供前瞻预览
    delivery_df = None
    if uploaded_file is not None:
        try:
            delivery_df = pd.read_excel(uploaded_file)
            delivery_df.columns = [str(c).strip() for c in delivery_df.columns]
            with st.expander("👀 展开查看上传发货单前 5 行预览", expanded=False):
                st.dataframe(delivery_df.head(5), use_container_width=True)
        except Exception as e:
            st.error(f"读取发货单失败: {e}")

    st.write("")
    
    # 开始处理按钮
    start_btn = st.button("🚀 开始拉取并计算数据", type="primary", use_container_width=True)

    if start_btn:
        if uploaded_file is None or delivery_df is None:
            st.warning("⚠️ 请先上传发货单 Excel 文件！")
            return

        with st.status("正在进行数据匹配与计算...", expanded=True) as status:
            try:
                # 步骤 1: 拉取采购单
                status.update(label="1/5 正在从 AirScript 云端拉取采购单数据...")
                purchase_df = fetch_purchase_from_airscript(webhook_url_input, token_input)
                purchase_df.columns = [str(c).strip() for c in purchase_df.columns]

                # 兼容品名/中文品名
                if "中文品名" not in purchase_df.columns and "品名" in purchase_df.columns:
                    purchase_df["中文品名"] = purchase_df["品名"]
                
                # 兼容单箱数量/箱规/PCS
                if "PCS" not in purchase_df.columns:
                    if "单箱数量" in purchase_df.columns:
                        purchase_df["PCS"] = purchase_df["单箱数量"]
                    elif "箱规" in purchase_df.columns:
                        purchase_df["PCS"] = purchase_df["箱规"]

                # 步骤 2: 校验表头字段
                status.update(label="2/5 正在校验表头字段规范...")
                required_purchase_cols = ["SKU", "中文品名", "PCS"]
                required_delivery_cols = ["SKU", "发货量", "货件编号"]

                for col in required_purchase_cols:
                    if col not in purchase_df.columns:
                        raise ValueError(f"云端采购单缺少必需列：【{col}】，请检查在线表格！")

                for col in required_delivery_cols:
                    if col not in delivery_df.columns:
                        raise ValueError(f"发货单缺少必需列：【{col}】，请检查 Excel 字段！")

                # 步骤 3: 关联匹配
                status.update(label="3/5 正在按 SKU 进行关联匹配...")
                purchase_df["SKU"] = purchase_df["SKU"].astype(str).str.strip()
                delivery_df["SKU"] = delivery_df["SKU"].astype(str).str.strip()

                purchase_mapping = purchase_df[["SKU", "中文品名", "PCS"]].drop_duplicates(subset=["SKU"])

                delivery_updated = pd.merge(
                    delivery_df,
                    purchase_mapping,
                    on="SKU",
                    how="left"
                )

                # 步骤 4: 数值计算
                status.update(label="4/5 正在计算总 PCS 及汇总报表...")
                delivery_updated["PCS"] = pd.to_numeric(delivery_updated["PCS"], errors="coerce").fillna(0)
                delivery_updated["发货量"] = pd.to_numeric(delivery_updated["发货量"], errors="coerce").fillna(0)
                delivery_updated["总PCS"] = delivery_updated["发货量"] * delivery_updated["PCS"]

                # 汇总统计
                summary_data = delivery_updated.groupby(["货件编号", "中文品名"]).agg({
                    "总PCS": "sum"
                }).reset_index()[["货件编号", "中文品名", "总PCS"]]

                status.update(label="✅ 数据处理完成！", state="complete")

                # 步骤 5: 结果展示与下载
                st.success("🎉 数据计算与汇总成功！请在下方预览并下载文件。")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                # 准备 Excel 字节流
                file1_bytes = to_excel_bytes(delivery_updated)
                file2_bytes = to_excel_bytes(summary_data)
                combined_bytes = to_combined_excel_bytes(delivery_updated, summary_data)

                # 下载区域
                st.markdown("### 📥 导出与下载")
                dl_col1, dl_col2, dl_col3 = st.columns(3)
                
                with dl_col1:
                    st.download_button(
                        label="📄 下载更新后发货单 (.xlsx)",
                        data=file1_bytes,
                        file_name=f"发货单_已更新_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with dl_col2:
                    st.download_button(
                        label="📊 下载货件汇总表 (.xlsx)",
                        data=file2_bytes,
                        file_name=f"货件编号_总PCS汇总表_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with dl_col3:
                    st.download_button(
                        label="📦 一键下载双表合一 (.xlsx)",
                        data=combined_bytes,
                        file_name=f"货件PCS全量导出_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                # 结果表格展示
                tab1, tab2 = st.tabs(["📄 已更新发货单明细", "📊 货件汇总表"])
                with tab1:
                    st.dataframe(delivery_updated, use_container_width=True)
                with tab2:
                    st.dataframe(summary_data, use_container_width=True)

            except Exception as e:
                status.update(label=f"❌ 处理出错：{str(e)}", state="error")
                st.error(f"处理失败: {str(e)}")

if __name__ == "__main__":
    main()