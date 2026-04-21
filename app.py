import streamlit as st
import pandas as pd
import io
import time

st.set_page_config(page_title="商品贴换标加工转换工具", layout="wide")
st.title("商品贴换标加工转换工具")

# ========== 第一步：上传库存表 ==========
st.header("第一步：上传《在库库存明细表》")
st.caption("系统自动筛选：仓库名称=DLM供应链深圳仓-SZ，且可用库存>0。")

inv_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls", "csv"], key="inv")

if inv_file:
    try:
        inventory_df = pd.read_excel(inv_file, dtype=str)
        # 数值列转换
        if "可用库存" in inventory_df.columns:
            inventory_df["可用库存"] = pd.to_numeric(inventory_df["可用库存"], errors="coerce").fillna(0)
        st.success(f"成功读取库存表，共加载 {len(inventory_df)} 行数据。")
        st.dataframe(inventory_df.head(), use_container_width=True)
    except Exception as e:
        st.error(f"读取文件失败：{e}")
        st.stop()
else:
    st.info("请上传库存明细表文件。")
    st.stop()

# ========== 法人主体映射 ==========
def map_legal_entity(val):
    if not val or str(val).strip() == "" or str(val) == "nan":
        return "未知主体"
    if "深圳市德拉姆供应链有限公司" in str(val):
        return "深圳市德拉姆供应链有限公司"
    return str(val).strip()

# ========== 第二步：录入转换需求 ==========
st.header("第二步：录入转换需求")
st.caption("支持类似 Excel 的操作：直接在表格中编辑，可粘贴批量数据。")

default_rows = 10
empty_df = pd.DataFrame({
    "SKU (必填)": [""] * default_rows,
    "目标 FNSKU (必填)": [""] * default_rows,
    "需求数量 (必填)": [None] * default_rows,
    "备注 (选填)": [""] * default_rows,
})

edited_df = st.data_editor(
    empty_df,
    num_rows="dynamic",
    use_container_width=True,
    key="req_table",
)

# ========== 执行匹配 ==========
if st.button("执行匹配并导出转换表", type="primary", use_container_width=True):

    # 解析需求数据
    requirements = []
    for _, row in edited_df.iterrows():
        sku = str(row.get("SKU (必填)", "")).strip()
        fnsku = str(row.get("目标 FNSKU (必填)", "")).strip()
        qty_raw = row.get("需求数量 (必填)")
        note = str(row.get("备注 (选填)", "")).strip()

        try:
            qty = int(float(qty_raw)) if qty_raw and str(qty_raw) != "nan" and str(qty_raw).strip() != "" else 0
        except (ValueError, TypeError):
            qty = 0

        if sku and fnsku and qty > 0:
            requirements.append({"sku": sku, "targetFnsku": fnsku, "qty": qty, "note": note})

    if not requirements:
        st.warning("请正确填写至少一条完整需求（SKU、FNSKU、数量必填且数量需大于0）！")
        st.stop()

    # 筛选有效库存
    valid_pool = inventory_df[
        (inventory_df["仓库名称"] == "DLM供应链深圳仓-SZ") &
        (inventory_df["可用库存"] > 0)
    ].copy()

    # 转为字典列表以便扣减（和原版逻辑一致）
    pool_records = valid_pool.to_dict("records")

    output_rows = []
    unfulfilled_rows = []
    summary_msgs = []

    for req in requirements:
        remain_qty = req["qty"]
        current_fulfilled = 0

        # 筛选该 SKU 的库存
        sku_inv = [r for r in pool_records if str(r.get("SKU", "")) == req["sku"]]

        # 优先匹配 FnSKU == 目标FNSKU 的排后面（先消耗不同FNSKU的，再消耗相同的）
        # 原版逻辑：FnSKU == targetFnsku 的排后面（return 1），不等于的排前面（return -1）
        sku_inv.sort(key=lambda x: 0 if str(x.get("FnSKU", "")) == req["targetFnsku"] else -1)

        for item in sku_inv:
            if remain_qty <= 0:
                break
            item_avail = float(item.get("可用库存", 0))
            if item_avail <= 0:
                continue

            deduct_qty = min(remain_qty, item_avail)
            remain_qty -= int(deduct_qty)
            current_fulfilled += int(deduct_qty)
            item["可用库存"] = item_avail - deduct_qty

            if str(item.get("FnSKU", "")) != req["targetFnsku"]:
                legal_entity = item.get("法人主体") or item.get("库存主体", "")
                output_rows.append({
                    "转换类型": "贴换标",
                    "法人主体": map_legal_entity(legal_entity),
                    "仓库": "DLM供应链深圳仓-SZ",
                    "SKU1": req["sku"],
                    "库区1": item.get("库区", ""),
                    "FNSKU1": item.get("FnSKU", ""),
                    "数量": int(deduct_qty),
                    "SKU2": req["sku"],
                    "FNSKU2": req["targetFnsku"],
                })

        if remain_qty > 0:
            legal_entity = ""
            if sku_inv:
                legal_entity = sku_inv[0].get("法人主体") or sku_inv[0].get("库存主体", "")
            unfulfilled_rows.append({
                "转换类型": "贴换标",
                "法人主体": map_legal_entity(legal_entity) if legal_entity else "无法确定主体",
                "仓库": "DLM供应链深圳仓-SZ",
                "SKU1": req["sku"],
                "库区1": sku_inv[0].get("库区", "") if sku_inv else "",
                "FNSKU1": "",
                "数量": remain_qty,
                "SKU2": req["sku"],
                "FNSKU2": req["targetFnsku"],
                "备注": f"不满足，总需{req['qty']}，只有{current_fulfilled}，缺{remain_qty}",
            })

        if remain_qty == 0:
            status = "✅ 满足"
        else:
            status = f"❌ 不满足，只有{current_fulfilled}，缺{remain_qty}"
        summary_msgs.append(f"需求 SKU: `{req['sku']}` | 需 {req['qty']} 个 -> **{status}**")

    # ========== 显示结果 ==========
    if not output_rows and not unfulfilled_rows:
        st.info("匹配完成！所有满足条件的库存 FnSKU 均与目标一致，无需贴标且无缺货，未生成单据。")
        st.stop()

    st.header("执行结果")
    for msg in summary_msgs:
        st.markdown(f"- {msg}")

    # ========== 生成 Excel ==========
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: 转换单
        if output_rows:
            df1 = pd.DataFrame(output_rows)
            # 插入注释行
            comment_row = pd.DataFrame([["第一行是注释"] + [""] * (len(df1.columns) - 1)], columns=df1.columns)
            header_row = pd.DataFrame([df1.columns.tolist()], columns=df1.columns)
            final_df1 = pd.concat([comment_row, header_row, df1], ignore_index=True)
            final_df1.to_excel(writer, sheet_name="转换单", index=False, header=False)
        else:
            pd.DataFrame([["第一行是注释"]]).to_excel(writer, sheet_name="转换单", index=False, header=False)

        # Sheet 2: 库存异常(缺货)
        if unfulfilled_rows:
            df2 = pd.DataFrame(unfulfilled_rows)
            comment_row2 = pd.DataFrame([["第一行是注释"] + [""] * (len(df2.columns) - 1)], columns=df2.columns)
            header_row2 = pd.DataFrame([df2.columns.tolist()], columns=df2.columns)
            final_df2 = pd.concat([comment_row2, header_row2, df2], ignore_index=True)
            final_df2.to_excel(writer, sheet_name="库存异常(缺货)", index=False, header=False)

    output.seek(0)

    # 下载按钮
    filename = f"贴换标转换单_{int(time.time())}.xlsx"
    st.download_button(
        label="下载转换表 Excel",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
