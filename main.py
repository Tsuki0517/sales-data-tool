import os
# 禁用streamlit版本与包元数据读取，彻底解决No package metadata报错
os.environ["STREAMLIT_DISABLE_VERSION_CHECK"] = "1"
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from streamlit_echarts import st_echarts

# 页面基础布局配置
st.set_page_config(page_title="出货备货分析系统（多文件合并+精准同比+单月筛选）", layout="wide")
st.title("📊 出货&备货数据分析系统 | 多文件合并导入 + 精准年月同比 + 单月单周筛选 + 多维度统计")

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
    required_cols = ["出货日期", "出货数量", "商品名称", "商品规格", "客户简称", "销售经理", "订单部门"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"❌ 表格缺失必填字段：{','.join(missing_cols)}，程序无法解析数据！")
        st.stop()

    # 日期与数值格式转换
    df["出货日期"] = pd.to_datetime(df["出货日期"], errors="coerce")
    df["出货数量"] = pd.to_numeric(df["出货数量"], errors="coerce")
    # 剔除空日期、空销量脏数据
    df = df.dropna(subset=["出货日期", "出货数量"])
    # 生成时间维度字段
    df["年月"] = df["出货日期"].dt.strftime("%Y-%m")
    df["年份"] = df["出货日期"].dt.year
    df["年周"] = df["出货日期"].dt.strftime("%Y-W%U")
    df["年+月"] = df["出货日期"].dt.to_period("M")
    df["年+周"] = df["出货日期"].dt.to_period("W")

    # 商品类别缺省填充
    if "商品类别" not in df.columns:
        df["商品类别"] = "未分类商品"
    # 客户类别缺省填充
    if "客户类别" not in df.columns:
        df["客户类别"] = "普通客户"

    return df

# 文件上传：支持多选多文件合并导入
def upload_file_area():
    st.subheader("📁 可按住Ctrl多选多个Excel文件，批量合并导入全部数据")
    uploaded_files = st.file_uploader("拖拽或多选 xls / xlsx 文件", type=["xls", "xlsx"], accept_multiple_files=True)
    df_list = []
    file_list = []
    if uploaded_files:
        for file in uploaded_files:
            file_list.append(file.name)
            stream = BytesIO(file.read())
            if file.name.endswith(".xls"):
                df_single = pd.read_excel(stream, engine="xlrd")
            else:
                df_single = pd.read_excel(stream, engine="openpyxl")
            df_clean = clean_raw_df(df_single)
            df_list.append(df_clean)
        merge_all = pd.concat(df_list, ignore_index=True)
        st.success(f"✅ 成功合并导入 {len(file_list)} 个文件：{', '.join(file_list)}")
        warn = []
        if "商品类别" not in merge_all.columns:
            warn.append("【商品类别】无该列，默认填充：未分类商品")
        if "客户类别" not in merge_all.columns:
            warn.append("【客户类别】无该列，默认填充：普通客户")
        if warn:
            st.warning("⚠️ 表格缺失字段自动补全：\n" + "\n".join(warn))
        return merge_all, "多文件合并数据集"
    else:
        st.info("请选择单个或多个Excel文件后继续分析")
        return None, None

# 生成商品汇总清单用于下拉筛选
def get_product_list(df):
    group_cols = ["商品类别", "商品名称", "商品规格", "订单部门", "客户类别"]
    product_df = df.groupby(group_cols)["出货数量"].sum().reset_index()
    product_df.rename(columns={"出货数量": "本年总销量"}, inplace=True)
    product_df = product_df.sort_values("本年总销量", ascending=False)
    product_df["完整型号"] = product_df["商品名称"] + " | " + product_df["商品规格"]
    return product_df

# 周期聚合+环比+【精准按年月匹配同比】
def get_period_agg_data(item_df, period_type="month"):
    if period_type == "month":
        agg_df = item_df.groupby("年+月")["出货数量"].sum().reset_index()
        agg_df["年月"] = agg_df["年+月"].astype(str)
        agg_df.rename(columns={"出货数量": "本期销量"}, inplace=True)
        # 环比：取上一行月份
        agg_df["上期销量"] = agg_df["本期销量"].shift(1)
        agg_df["月环比"] = agg_df["本期销量"].pct_change(1).apply(format_pct)

        # 构建年月-销量映射表，精准匹配去年同月
        ym_sale_map = dict(zip(agg_df["年月"], agg_df["本期销量"]))
        def get_last_year_ym(ym_str):
            year_part = int(ym_str[:4])
            month_part = ym_str[5:]
            target_ym = f"{year_part - 1}-{month_part}"
            return ym_sale_map.get(target_ym, None)
        agg_df["去年同期销量"] = agg_df["年月"].apply(get_last_year_ym)

        # 同比计算
        def calc_tongbi(row):
            curr = row["本期销量"]
            last = row["去年同期销量"]
            if last is None or last == 0:
                return None
            return (curr - last) / last
        agg_df["同比"] = agg_df.apply(calc_tongbi, axis=1).apply(format_pct)
        x_axis_col = "年月"
    else:
        # 按周统计无同比，仅环比
        agg_df = item_df.groupby("年+周")["出货数量"].sum().reset_index()
        agg_df["年周"] = agg_df["年+周"].astype(str)
        agg_df.rename(columns={"出货数量": "本期销量"}, inplace=True)
        agg_df["上周销量"] = agg_df["本期销量"].shift(1)
        agg_df["周环比"] = agg_df["本期销量"].pct_change(1).apply(format_pct)
        agg_df["同比"] = "无"
        x_axis_col = "年周"
    return agg_df, x_axis_col

# 核心统计计算函数
def get_single_product_data(df, target_name, target_spec, target_category, target_dept, target_cust_type, filter_time="全部", period_type="month"):
    # 基础商品筛选
    base_mask = (df["商品名称"] == target_name) & (df["商品规格"] == target_spec) & (df["商品类别"] == target_category)
    data_base = df[base_mask].copy()
    # 订单部门筛选
    if target_dept != "全部":
        data_base = data_base[data_base["订单部门"] == target_dept]
    # 客户类别筛选
    if target_cust_type != "全部":
        data_base = data_base[data_base["客户类别"] == target_cust_type]
    if data_base.empty:
        return None

    # 全量数据用于上方趋势同比表格
    data_full = data_base.copy()
    # 单周期切片数据用于下方明细图表
    data_slice = data_base.copy()
    if filter_time != "全部":
        if period_type == "month":
            data_slice = data_slice[data_slice["年月"] == filter_time]
        else:
            data_slice = data_slice[data_slice["年周"] == filter_time]

    # 总累计出货
    total_total = data_full["出货数量"].sum()
    # 备货：近90天总销量 / 3 * 1.2
    ninety_day_ago = datetime.now() - timedelta(days=90)
    sale_90 = data_full[data_full["出货日期"] >= ninety_day_ago]["出货数量"].sum()
    avg_month = sale_90 / 3
    stock_suggest = int(avg_month * 1.2)

    # 全周期趋势与同比环比
    period_table, x_key = get_period_agg_data(data_full, period_type)
    x_list = period_table[x_key].tolist()
    y_list = period_table["本期销量"].tolist()

    # 客户维度统计（受单周期筛选）
    df_cust = data_slice.groupby("客户简称")["出货数量"].sum().reset_index()
    df_cust.rename(columns={"出货数量": "采购总件数"}, inplace=True)
    df_cust = df_cust.sort_values("采购总件数", ascending=False)
    cust_x = df_cust["客户简称"].tolist()
    cust_y = df_cust["采购总件数"].tolist()

    # 业务员维度统计
    df_sale = data_slice.groupby("销售经理")["出货数量"].sum().reset_index()
    df_sale.rename(columns={"出货数量": "销售总件数"}, inplace=True)
    df_sale = df_sale.sort_values("销售总件数", ascending=False)
    sale_x = df_sale["销售经理"].tolist()
    sale_y = df_sale["销售总件数"].tolist()

    # 订单部门维度统计
    df_dept = data_slice.groupby("订单部门")["出货数量"].sum().reset_index()
    df_dept.rename(columns={"出货数量": "部门销售总件数"}, inplace=True)
    df_dept = df_dept.sort_values("部门销售总件数", ascending=False)
    dept_x = df_dept["订单部门"].tolist()
    dept_y = df_dept["部门销售总件数"].tolist()

    # 客户类别汇总
    df_cust_type = data_slice.groupby("客户类别")["出货数量"].sum().reset_index()
    df_cust_type.rename(columns={"出货数量": "类别销量"}, inplace=True)

    return {
        "meta": {
            "总累计出货": int(total_total),
            "近三月总出货": int(sale_90),
            "推荐备货量": stock_suggest,
            "筛选范围": filter_time,
            "统计维度": "按月统计" if period_type == "month" else "按周统计",
            "商品大类": target_category,
            "所选部门": target_dept,
            "客户分类": target_cust_type
        },
        "trend_table": period_table,
        "trend_x": x_list,
        "trend_y": y_list,
        "cust_data": df_cust,
        "cust_x": cust_x,
        "cust_y": cust_y,
        "seller_data": df_sale,
        "seller_x": sale_x,
        "seller_y": sale_y,
        "dept_data": df_dept,
        "dept_x": dept_x,
        "dept_y": dept_y,
        "cust_type_data": df_cust_type
    }

# 绘图工具
def draw_line(x_data, y_data, title_text):
    opt = {
        "tooltip": {"trigger": "axis"},
        "xAxis": {"type": "category", "data": x_data},
        "yAxis": {"type": "value"},
        "series": [{"name": title_text, "type": "line", "data": y_data, "smooth": True}]
    }
    st_echarts(options=opt, height="360px")

def draw_bar(x_data, y_data, title_text):
    opt = {
        "tooltip": {"trigger": "axis"},
        "xAxis": {"type": "category", "data": x_data, "axisLabel": {"rotate": 35}},
        "yAxis": {"type": "value"},
        "series": [{"name": title_text, "type": "bar", "data": y_data}]
    }
    st_echarts(options=opt, height="360px")

# 主程序入口
if __name__ == "__main__":
    df_main, file_info = upload_file_area()
    if df_main is None:
        st.stop()

    with st.sidebar:
        st.header("🔍 多级筛选控制面板")
        product_pool = get_product_list(df_main)

        # 1 商品大类
        cate_options = sorted(product_pool["商品类别"].unique().tolist())
        select_cate = st.selectbox("① 选择商品大类", cate_options)
        filter_1 = product_pool[product_pool["商品类别"] == select_cate]

        # 2 订单部门筛选（带全部）
        dept_options = ["全部"] + sorted(filter_1["订单部门"].unique().tolist())
        select_dept = st.selectbox("② 选择订单部门", dept_options)
        if select_dept == "全部":
            filter_2 = filter_1
        else:
            filter_2 = filter_1[filter_1["订单部门"] == select_dept]

        # 3 客户类别筛选（带全部）
        cust_options = ["全部"] + sorted(filter_2["客户类别"].unique().tolist())
        select_cust = st.selectbox("③ 选择客户类别", cust_options)
        if select_cust == "全部":
            filter_3 = filter_2
        else:
            filter_3 = filter_2[filter_2["客户类别"] == select_cust]

        # 4 商品型号下拉，移除搜索框
        if len(filter_3) == 0:
            st.warning("当前筛选无匹配商品，请更换条件")
            st.stop()
        model_list = filter_3["完整型号"].unique().tolist()
        select_model = st.selectbox("④ 选择具体商品型号", model_list)
        name_split, spec_split = select_model.split(" | ")

        # 5 统计维度：按月 / 按周
        mode = st.radio("⑤ 统计分析维度", ["按月统计", "按周统计"])
        mode_flag = "month" if mode == "按月统计" else "week"

        # 6 单周期筛选下拉（可选单独某月/某周查看明细）
        if mode_flag == "month":
            time_options = ["全部"] + sorted(df_main["年月"].unique().tolist())
        else:
            time_options = ["全部"] + sorted(df_main["年周"].unique().tolist())
        select_time_slice = st.selectbox("⑥ 单独筛选单月/单周明细（选全部则查看全周期）", time_options)

        # 缓存刷新按钮
        if st.button("🔄 清空缓存刷新数据（更换文件必须点击）"):
            st.cache_data.clear()
            st.rerun()

    # 执行计算
    result = get_single_product_data(
        df=df_main,
        target_name=name_split,
        target_spec=spec_split,
        target_category=select_cate,
        target_dept=select_dept,
        target_cust_type=select_cust,
        filter_time=select_time_slice,
        period_type=mode_flag
    )
    if result is None:
        st.error("❌ 当前筛选组合无对应出货数据，请调整筛选条件")
        st.stop()

    # 顶部状态栏说明
    info_str = (
        f"📅 {result['meta']['统计维度']} | 明细筛选范围：{result['meta']['筛选范围']}\n"
        f"📦 商品大类：{result['meta']['商品大类']} | 📋 订单部门：{result['meta']['所选部门']} | 👥 客户分类：{result['meta']['客户分类']}"
    )
    st.info(info_str)
    st.divider()

    # 顶部三大指标卡片
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("该筛选条件累计总出货", f"{result['meta']['总累计出货']} 件")
    with col_b:
        st.metric("近三个月合计出货", f"{result['meta']['近三月总出货']} 件")
    with col_c:
        st.metric("系统推荐备货数量", f"{result['meta']['推荐备货量']} 件")
    st.divider()

    # 全周期趋势折线+同比环比总表（不受单月筛选影响，保证同比完整）
    st.subheader("📈 全周期销量趋势 & 环比同比明细（全量数据，同比以此表为准）")
    draw_line(result["trend_x"], result["trend_y"], "月度出货量")
    st.dataframe(result["trend_table"], use_container_width=True, hide_index=True)
    st.divider()

    st.subheader(f"📊 下方图表与明细表格仅展示【{select_time_slice}】范围内切片数据")
    # 客户分布 & 订单部门分布
    col_left, col_right = st.columns(2)
    with col_left:
        st.subheader("🏢 客户采购数量分布")
        if len(result["cust_x"]) == 0:
            st.warning("该时间段无出货记录")
        else:
            draw_bar(result["cust_x"], result["cust_y"], "采购件数")
        st.dataframe(result["cust_data"], height=300, use_container_width=True, hide_index=True)
    with col_right:
        st.subheader("📋 各订单部门该产品销量")
        if len(result["dept_x"]) == 0:
            st.warning("该时间段无出货记录")
        else:
            draw_bar(result["dept_x"], result["dept_y"], "部门销售件数")
        st.dataframe(result["dept_data"], height=300, use_container_width=True, hide_index=True)
    st.divider()

    # 业务员业绩 & 客户类别汇总
    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("👤 销售经理个人业绩分布")
        if len(result["seller_x"]) == 0:
            st.warning("该时间段无出货记录")
        else:
            draw_bar(result["seller_x"], result["seller_y"], "销售件数")
    with col_r:
        st.subheader("👥 客户类别销量汇总")
        st.dataframe(result["cust_type_data"], height=320, use_container_width=True, hide_index=True)

    # 业务员明细清单
    st.subheader("📋 业务员销售明细清单")
    st.dataframe(result["seller_data"], use_container_width=True, hide_index=True)