import os
os.environ["STREAMLIT_DISABLE_VERSION_CHECK"] = "1"
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
os.environ["STREAMLIT_SERVER_METRICS_ENABLED"] = "false"
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from streamlit_echarts import st_echarts

# 全局备货安全系数
SAFE_MULTIPLE = 1.5
# 页面基础布局配置
st.set_page_config(page_title="出货备货分析系统（周/月环比+同比+订单部门统计）", layout="wide")
st.title("📊 出货&备货数据分析系统 | 环比同比 + 客户 + 业务员 + 订单部门多维度统计")

# 百分比格式化工具（空值/除零统一显示“无”）
def format_pct(val):
    if pd.isna(val) or val == float('inf') or val == -float('inf') or val is None:
        return "无"
    return f"{round(val * 100, 2)}%"

# 数据清洗&标准化
def clean_raw_df(raw_df):
    df = raw_df.copy()
    # 清除列名前后空格
    df.columns = [c.strip() for c in df.columns]
    # 强制校验必填核心列
    required_cols = ["出货日期", "出货数量", "商品名称", "商品规格", "客户简称", "销售经理"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"❌ 表格缺失必填字段：{','.join(missing_cols)}，程序无法解析数据！")
        st.stop()

    # 日期与数值格式转换
    df["出货日期"] = pd.to_datetime(df["出货日期"], errors="coerce")
    df["出货数量"] = pd.to_numeric(df["出货数量"], errors="coerce")
    # 剔除脏数据
    df = df.dropna(subset=["出货日期", "出货数量"])
    # 生成多维度时间字段
    df["年月"] = df["出货日期"].dt.strftime("%Y-%m")
    df["年份"] = df["出货日期"].dt.year
    df["年周"] = df["出货日期"].dt.strftime("%Y-W%U")  # 自然周标识
    df["年+月"] = df["出货日期"].dt.to_period("M")
    df["年+周"] = df["出货日期"].dt.to_period("W")

    # 适配【订单部门】列，无该列则统一填充为「未填写订单部门」
    if "订单部门" not in df.columns:
        df["订单部门"] = "未填写订单部门"
    return df

def upload_file_area():
    st.subheader("📁 上传出货表格（支持多文件同时上传）")
    uploaded_files = st.file_uploader("将Excel文件拖入下方框内（可多选）", type=["xls", "xlsx"], accept_multiple_files=True)
    
    if uploaded_files is not None and len(uploaded_files) > 0:
        dfs = []
        file_names = []
        has_missing_dept = False
        
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            file_names.append(file_name)
            st.success(f"✅ 已成功载入文件：{file_name}")
            file_stream = BytesIO(uploaded_file.read())
            if file_name.endswith(".xls"):
                raw_data = pd.read_excel(file_stream, engine="xlrd")
            else:
                raw_data = pd.read_excel(file_stream, engine="openpyxl")
            if "订单部门" not in raw_data.columns:
                has_missing_dept = True
            dfs.append(raw_data)
        
        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            clean_df = clean_raw_df(combined_df)
            if has_missing_dept:
                st.warning("⚠️ 部分文件无【订单部门】列，对应数据将归类为「未填写订单部门」")
            return clean_df, ", ".join(file_names)
        else:
            return None, None
    else:
        st.info("请上传出货明细Excel文件后继续操作（支持多选）")
        return None, None

# 汇总全部商品清单
def get_product_list(df):
    product_df = df.groupby(["商品类别", "商品名称", "商品规格"])["出货数量"].sum().reset_index()
    product_df.rename(columns={"出货数量": "本年总销量"}, inplace=True)
    product_df = product_df.sort_values("本年总销量", ascending=False)
    product_df["完整型号"] = product_df["商品名称"] + " | " + product_df["商品规格"]
    return product_df

def get_period_agg_data(item_df, period_type="month"):
    if period_type == "month":
        agg_df = item_df.groupby("年+月")["出货数量"].sum().reset_index()
        agg_df["年月"] = agg_df["年+月"].astype(str)
        agg_df.rename(columns={"出货数量": "本期销量"}, inplace=True)
        
        agg_df["上期销量"] = agg_df["本期销量"].shift(1)
        agg_df["月环比"] = agg_df["本期销量"].pct_change(1).apply(format_pct)
        
        agg_df["年份"] = agg_df["年月"].str[:4].astype(int)
        agg_df["月份"] = agg_df["年月"].str[-2:].astype(int)
        
        year_month_map = agg_df.set_index(["年份", "月份"])["本期销量"].to_dict()
        
        agg_df["去年同期销量"] = agg_df.apply(lambda row: year_month_map.get((row["年份"] - 1, row["月份"]), None), axis=1)
        
        def calc_yoy(row):
            if row["去年同期销量"] is None or row["去年同期销量"] == 0:
                return "无"
            return format_pct((row["本期销量"] - row["去年同期销量"]) / row["去年同期销量"])
        
        agg_df["同比"] = agg_df.apply(calc_yoy, axis=1)
        x_col = "年月"
    else:
        agg_df = item_df.groupby("年+周")["出货数量"].sum().reset_index()
        agg_df["年周"] = agg_df["年+周"].astype(str)
        agg_df.rename(columns={"出货数量": "本期销量"}, inplace=True)
        agg_df["上周销量"] = agg_df["本期销量"].shift(1)
        agg_df["周环比"] = agg_df["本期销量"].pct_change(1).apply(format_pct)
        
        agg_df["年份"] = agg_df["年周"].str[:4].astype(int)
        agg_df["周数"] = agg_df["年周"].str[-2:].astype(int)
        
        year_week_map = agg_df.set_index(["年份", "周数"])["本期销量"].to_dict()
        
        agg_df["去年同期销量"] = agg_df.apply(lambda row: year_week_map.get((row["年份"] - 1, row["周数"]), None), axis=1)
        
        def calc_yoy(row):
            if row["去年同期销量"] is None or row["去年同期销量"] == 0:
                return "无"
            return format_pct((row["本期销量"] - row["去年同期销量"]) / row["去年同期销量"])
        
        agg_df["同比"] = agg_df.apply(calc_yoy, axis=1)
        x_col = "年周"
    return agg_df, x_col

# 提取选中单品全维度数据
def get_single_product_data(df, target_name, target_spec, filter_time="全部", period_type="month"):
    item_data = df[(df["商品名称"] == target_name) & (df["商品规格"] == target_spec)]
    if item_data.empty:
        return None

    # 时间筛选
    filter_df = item_data.copy()
    if filter_time != "全部":
        if period_type == "month":
            filter_df = item_data[item_data["年月"] == filter_time]
        else:
            filter_df = item_data[item_data["年周"] == filter_time]

    # 总指标
    total_all = item_data["出货数量"].sum()
    cutoff_day = datetime.now() - timedelta(days=30)
    last30 = item_data[item_data["出货日期"] >= cutoff_day]["出货数量"].sum()
    suggest_stock = int(last30 * SAFE_MULTIPLE)

    # 周期聚合+环比同比
    period_table, x_key = get_period_agg_data(item_data, period_type)
    x_list = period_table[x_key].tolist()
    y_list = period_table["本期销量"].tolist()

    # 1.客户统计
    cust_df = filter_df.groupby("客户简称")["出货数量"].sum().reset_index()
    cust_df.rename(columns={"出货数量": "采购总件数"}, inplace=True)
    cust_df = cust_df.sort_values("采购总件数", ascending=False)
    cust_x = cust_df["客户简称"].tolist()
    cust_y = cust_df["采购总件数"].tolist()

    # 2.业务员统计
    seller_df = filter_df.groupby("销售经理")["出货数量"].sum().reset_index()
    seller_df.rename(columns={"出货数量": "销售总件数"}, inplace=True)
    seller_df = seller_df.sort_values("销售总件数", ascending=False)
    seller_x = seller_df["销售经理"].tolist()
    seller_y = seller_df["销售总件数"].tolist()

    # 3.订单部门维度统计（已全部修改为订单部门）
    dept_df = filter_df.groupby("订单部门")["出货数量"].sum().reset_index()
    dept_df.rename(columns={"出货数量": "部门销售总件数"}, inplace=True)
    dept_df = dept_df.sort_values("部门销售总件数", ascending=False)
    dept_x = dept_df["订单部门"].tolist()
    dept_y = dept_df["部门销售总件数"].tolist()

    return {
        "index": {
            "总销量": int(total_all),
            "近30天": int(last30),
            "备货建议": suggest_stock,
            "筛选范围": filter_time,
            "统计维度": "按月统计" if period_type == "month" else "按周统计"
        },
        "period_table": period_table,
        "x_data": x_list,
        "y_data": y_list,
        "cust_table": cust_df,
        "cust_x": cust_x,
        "cust_y": cust_y,
        "seller_table": seller_df,
        "seller_x": seller_x,
        "seller_y": seller_y,
        "dept_table": dept_df,
        "dept_x": dept_x,
        "dept_y": dept_y
    }

# 绘图函数
def draw_trend_line(x_list, y_list, title_name):
    opt = {
        "tooltip": {"trigger": "axis"},
        "xAxis": {"type": "category", "data": x_list},
        "yAxis": {"type": "value"},
        "series": [{"name": title_name, "type": "line", "data": y_list, "smooth": True}]
    }
    st_echarts(options=opt, height="360px")

def draw_bar(x_list, y_list, bar_name):
    opt = {
        "tooltip": {"trigger": "axis"},
        "xAxis": {"type": "category", "data": x_list, "axisLabel": {"rotate": 35}},
        "yAxis": {"type": "value"},
        "series": [{"name": bar_name, "type": "bar", "data": y_list}]
    }
    st_echarts(options=opt, height="360px")

# 主程序
if __name__ == "__main__":
    df_data, file_name = upload_file_area()
    if df_data is None:
        st.stop()

    # 侧边栏
    with st.sidebar:
        st.header("🔍 筛选控制面板")
        all_product = get_product_list(df_data)
        search_key = st.text_input("商品名称关键词检索", value="")
        if search_key.strip():
            filter_prod = all_product[all_product["商品名称"].str.contains(search_key, case=False, na=False)]
            st.success(f"匹配 {len(filter_prod)} 个型号")
        else:
            filter_prod = all_product
        selected_item = st.selectbox("选择商品型号", filter_prod["完整型号"].tolist())
        item_name, item_spec = selected_item.split(" | ")

        # 选择统计周期
        period_mode = st.radio("统计分析维度", ["按月统计", "按周统计"])
        p_type = "month" if period_mode == "按月统计" else "week"

        # 可选时间范围
        if p_type == "month":
            time_list = ["全部"] + sorted(df_data["年月"].unique().tolist())
        else:
            time_list = ["全部"] + sorted(df_data["年周"].unique().tolist())
        select_time = st.selectbox("筛选单周期查看", time_list)

        if st.button("🔄 刷新数据缓存"):
            st.cache_data.clear()
            st.rerun()

    # 计算结果
    res = get_single_product_data(df_data, item_name, item_spec, select_time, p_type)
    if res is None:
        st.error("未查询到该商品出货数据")
        st.stop()

    st.info(f"📅 当前：{res['index']['统计维度']} | 筛选范围：{res['index']['筛选范围']}")
    available_years = sorted(df_data["年份"].unique())
    if len(available_years) < 2:
        st.warning(f"⚠️ 当前数据仅包含 {available_years} 年份，同比数据无法计算。请上传包含至少两年数据的文件以查看同比分析。")
    st.divider()

    # 顶部指标卡片
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("该商品累计总出货", f"{res['index']['总销量']} 件")
    with c2:
        st.metric("近30天实际销量", f"{res['index']['近30天']} 件")
    with c3:
        st.metric("系统推荐备货量", f"{res['index']['备货建议']} 件")
    st.divider()

    # 周期趋势+环比同比表格
    st.subheader(f"📈 {period_mode}销量趋势 & 环比/同比明细")
    draw_trend_line(res["x_data"], res["y_data"], "出货销量")
    st.dataframe(res["period_table"], use_container_width=True, hide_index=True)
    st.divider()

    # 第一行：客户 + 订单部门
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🏢 客户采购数量分布")
        if len(res["cust_x"]) == 0:
            st.warning("该周期无出货记录")
        else:
            draw_bar(res["cust_x"], res["cust_y"], "采购件数")
        st.dataframe(res["cust_table"], height=300, use_container_width=True, hide_index=True)
    with col2:
        st.subheader("📋 各订单部门该产品销量")
        if len(res["dept_x"]) == 0:
            st.warning("该周期无出货记录")
        else:
            draw_bar(res["dept_x"], res["dept_y"], "部门销售件数")
        st.dataframe(res["dept_table"], height=300, use_container_width=True, hide_index=True)
    st.divider()

    # 第二行：业务员业绩
    col3, col4 = st.columns(2)
    with col3:
        st.subheader("👤 销售经理个人业绩分布")
        if len(res["seller_x"]) == 0:
            st.warning("该周期无出货记录")
        else:
            draw_bar(res["seller_x"], res["seller_y"], "销售件数")
    with col4:
        st.subheader("业务员销售明细清单")
        st.dataframe(res["seller_table"], height=320, use_container_width=True, hide_index=True)