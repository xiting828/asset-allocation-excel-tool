"""Streamlit 网页入口。"""

from __future__ import annotations

import csv
import hashlib

import pandas as pd
import streamlit as st

from table_tool import (
    build_excel,
    download_filename,
    find_asset_category_column,
    parse_pasted_table,
    preview_parts,
    TableLimitError,
)


st.set_page_config(page_title="粘贴表格转 Excel", page_icon="📊", layout="wide")

st.title("粘贴表格转 Excel")
st.caption("把 Excel、网页或聊天软件中的表格粘贴到下方，预览清理结果，然后下载排版好的 Excel 文件。")
st.info("您粘贴的数据仅用于本次 Excel 文件生成，不会被主动保存。请勿输入银行卡号、身份证号等高度敏感信息。")

with st.expander("使用说明", expanded=True):
    st.markdown(
        """
1. 复制使用 **Tab 分列、换行分行** 的表格数据并粘贴。
2. 程序会把所有 `*` 替换为 `0`，仅删除表格外围的全空白行和列。
3. 确认预览后，点击下方按钮生成并下载 `.xlsx` 文件。

带前导零的编号会优先按文本保存；各行列数不一致时会自动在末尾补空单元格。“资产大类”列中的“固定收益类”和“权益类”会分别使用淡绿色和淡蓝色区分。
"""
    )

first_row_is_header = st.checkbox("第一行是表头", value=True)
pasted_text = st.text_area(
    "粘贴表格数据",
    height=280,
    placeholder="例如：\n姓名\t学号\t成绩\n张三\t00123\t95\n李四\t00124\t*",
)

parse_error: str | None = None
try:
    parsed = parse_pasted_table(pasted_text)
except TableLimitError as exc:
    parsed = parse_pasted_table("")
    parse_error = str(exc)
except (csv.Error, UnicodeError):
    parsed = parse_pasted_table("")
    parse_error = "无法识别粘贴的数据，请确认内容使用 Tab 分列、换行分行后重试。"

if parse_error:
    st.error(parse_error)
input_signature = hashlib.sha256(
    f"{first_row_is_header}\0{pasted_text}".encode("utf-8")
).hexdigest()

if not parsed.is_empty:
    if parsed.inconsistent_columns:
        counts = "、".join(str(count) for count in sorted(set(parsed.original_column_counts)))
        st.warning(f"检测到各行列数不一致（出现了 {counts} 列）；已在较短行末尾补充空单元格。")

    if find_asset_category_column(parsed.rows, first_row_is_header) is None:
        st.info("未找到“资产大类”列，将继续正常生成 Excel，但不会自动添加资产类别背景色。")

    headers, data_rows = preview_parts(parsed.rows, first_row_is_header)
    st.subheader("处理后预览")
    st.caption(f"共 {len(data_rows)} 行数据，{parsed.column_count} 列")
    preview_frame = pd.DataFrame(data_rows, columns=headers, dtype=object)
    centered_preview = preview_frame.style.set_properties(
        **{"text-align": "center", "vertical-align": "middle"}
    ).set_table_styles(
        [
            {"selector": "th", "props": [("text-align", "center"), ("vertical-align", "middle")]},
            {"selector": "td", "props": [("text-align", "center"), ("vertical-align", "middle")]},
        ]
    )
    centered_columns = {
        header: st.column_config.TextColumn(header, alignment="center") for header in headers
    }
    st.dataframe(
        centered_preview,
        column_config=centered_columns,
        width="stretch",
        hide_index=True,
    )

st.divider()


def clear_generated_excel() -> None:
    """下载触发后清除会话内的 Excel 字节，不在服务器会话中继续保留。"""

    st.session_state.pop("excel_data", None)
    st.session_state.pop("excel_filename", None)
    st.session_state.pop("excel_input_signature", None)

if st.button("生成并下载 Excel", type="primary", width="stretch"):
    if parse_error:
        st.error("请先调整输入数据，再生成 Excel。")
    elif parsed.is_empty:
        st.error("还没有可处理的数据。请先在上方粘贴表格内容。")
    else:
        try:
            excel_data = build_excel(parsed.rows, first_row_is_header)
            st.session_state["excel_data"] = excel_data.getvalue()
            st.session_state["excel_filename"] = download_filename()
            st.session_state["excel_input_signature"] = input_signature
            st.success("Excel 已生成，请点击下方按钮下载。")
        except (ValueError, MemoryError):
            st.error("生成 Excel 时出现问题，请减少数据量或检查表格格式后重试。")
        except Exception:
            st.error("暂时无法生成 Excel，请稍后重试。")

if (
    "excel_data" in st.session_state
    and st.session_state.get("excel_input_signature") == input_signature
):
    st.download_button(
        "下载整理后的 Excel",
        data=st.session_state["excel_data"],
        file_name=st.session_state["excel_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
        on_click=clear_generated_excel,
    )
