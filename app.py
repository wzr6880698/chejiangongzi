import openpyxl
from openpyxl import load_workbook
import os
from datetime import datetime, timedelta
import re
import streamlit as st
from io import BytesIO
import tempfile
import pandas as pd

# ============================
# 工具类
# ============================
class DateParser:
    @staticmethod
    def parse(date_str):
        if not date_str:
            return None
        if isinstance(date_str, (int, float)):
            return DateParser._parse_excel_number(date_str)
        date_str = str(date_str).strip()
        formats = [
            (r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", "%Y-%m-%d"),
            (r"\d{4}年\d{1,2}月\d{1,2}日", "%Y年%m月%d日"),
            (r"\d{2}年\d{1,2}月\d{1,2}日", "%y年%m月%d日"),
            (r"\d{1,2}月\d{1,2}日", "%m月%d日"),
            (r"\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}", "%Y-%m-%d")
        ]
        for pattern, fmt in formats:
            if re.match(pattern, date_str):
                try:
                    dt = datetime.strptime(date_str.split()[0] if " " in date_str else date_str, fmt)
                    return dt.strftime("%Y/%m/%d")
                except ValueError:
                    continue
        return None

    @staticmethod
    def _parse_excel_number(num):
        try:
            base_date = datetime(1899, 12, 30)
            delta = timedelta(days=int(num))
            return (base_date + delta).strftime("%Y/%m/%d")
        except (ValueError, TypeError):
            return None


class DataValidator:
    @staticmethod
    def is_valid_name(name):
        if not name or not isinstance(name, str):
            return False
        name = name.strip()
        return (name and len(name) >= 2 and
                name not in ["姓名", "合计", "序号", None, "日期", "车间生产日报表", "生产日报表"])

    @staticmethod
    def is_valid_number(value):
        try:
            float(value)
            return True
        except (TypeError, ValueError):
            return False

    @staticmethod
    def validate_record(record):
        required_fields = ["日期", "姓名", "产品名称"]
        for field in required_fields:
            if not record.get(field):
                return False
        return True


# ============================
# 基础提取器（支持多表头块）
# ============================
class WorkshopDataExtractor:
    def __init__(self, sheet_name):
        self.sheet_name = sheet_name

    def extract(self, ws, data_list, manual_config=None):
        if manual_config:
            self._extract_with_manual_config(ws, data_list, manual_config)
        else:
            # 动态解析多表头块
            success = self._try_dynamic_extract(ws, data_list)
            if not success:
                self._fallback_extract(ws, data_list)

    def _try_dynamic_extract(self, ws, data_list):
        """扫描整个工作表，识别所有表头块并提取数据"""
        header_blocks = self._find_header_blocks(ws)
        if not header_blocks:
            return False

        total_records = 0
        # 对每个块处理
        for block in header_blocks:
            # 提取该块的数据
            records = self._extract_block_data(ws, block, data_list)
            total_records += records
        return total_records > 0

    def _find_header_blocks(self, ws):
        """
        查找所有表头块。
        返回列表，每个元素为dict:
        {
            'header_row': 行号（1-based）,
            'date_col': 日期列索引（1-based）或None,
            'name_col': 姓名列索引,
            'batch_col': 批次号列索引或None,
            'product_col': 产品名列索引或None,  # 如果产品名每块固定，可从上方解析
            'blocks': [  # 产品块列表
                {'q_col': 数量列, 'price_col': 单价列, 'amount_col': 金额列, 'note_col': 备注列, 'batch': 批次号, 'product': 产品名称}
            ]
        }
        """
        blocks = []
        max_row = ws.max_row
        max_col = ws.max_column

        # 第一步：找出所有包含"数量"的单元格（行，列）
        quantity_cells = []
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                if cell.value and isinstance(cell.value, str):
                    val = cell.value.strip()
                    if val == "数量":
                        quantity_cells.append((r, c))

        if not quantity_cells:
            return []

        # 按行分组
        from collections import defaultdict
        rows_with_q = defaultdict(list)
        for r, c in quantity_cells:
            rows_with_q[r].append(c)

        # 对每个出现"数量"的行，构造一个表头块
        for header_row, q_cols in sorted(rows_with_q.items()):
            # 查找该行上的其他关键字段（日期、姓名、批次号、产品名称）
            # 注意：这些字段可能不在同一行，向上搜索
            date_col = None
            name_col = None
            # 先查看该行是否有"日期"、"姓名"等关键词
            for c in range(1, max_col + 1):
                cell = ws.cell(header_row, c)
                if cell.value and isinstance(cell.value, str):
                    val = cell.value.strip()
                    if val == "日期":
                        date_col = c
                    elif val == "姓名":
                        name_col = c
            # 若姓名列未找到，默认B列（常见）
            if name_col is None:
                # 尝试从上一行找"姓名"
                for r in range(header_row - 1, max(1, header_row - 5), -1):
                    for c in range(1, max_col + 1):
                        cell = ws.cell(r, c)
                        if cell.value and isinstance(cell.value, str) and cell.value.strip() == "姓名":
                            name_col = c
                            break
                    if name_col:
                        break
            if name_col is None:
                name_col = 2  # 默认B列

            # 处理每个数量列，生成产品块
            product_blocks = []
            for q_col in sorted(q_cols):
                # 向上查找批次号和产品名称（可能在当前行或上方行）
                batch = "0"
                product = f"产品{q_col}"
                # 向上扫描最多5行
                for r in range(header_row, max(1, header_row - 5), -1):
                    cell_val = ws.cell(r, q_col).value
                    if cell_val and isinstance(cell_val, str):
                        val = cell_val.strip()
                        if val in ["批次号", "批号"]:
                            # 批次号值在右侧列
                            batch_cell = ws.cell(r, q_col + 1)
                            if batch_cell.value is not None:
                                batch = str(batch_cell.value).strip()
                        elif val in ["产品名称", "品名"]:
                            prod_cell = ws.cell(r, q_col + 1)
                            if prod_cell.value is not None:
                                product = str(prod_cell.value).strip()
                # 确定其他字段列（默认紧随其后）
                price_col = q_col + 1
                amount_col = q_col + 2
                note_col = q_col + 3
                # 但也要检查是否确实有"单价"等标签，如果没有，可能列顺序不同？暂信任顺序
                product_blocks.append({
                    'q_col': q_col,
                    'price_col': price_col,
                    'amount_col': amount_col,
                    'note_col': note_col,
                    'batch': batch,
                    'product': product
                })

            blocks.append({
                'header_row': header_row,
                'date_col': date_col,
                'name_col': name_col,
                'product_blocks': product_blocks
            })

        return blocks

    def _extract_block_data(self, ws, block, data_list):
        """从给定块中提取数据，添加到data_list"""
        header_row = block['header_row']
        name_col = block['name_col']
        date_col = block.get('date_col')
        product_blocks = block['product_blocks']

        # 确定数据起始行（表头下一行）
        start_row = header_row + 1
        # 查找下一个表头行（即下一个包含"数量"的行），作为结束行
        end_row = ws.max_row
        # 可以简单扫描后续行，若某行有"数量"且行号大于header_row，则作为结束
        for r in range(start_row, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if cell.value and isinstance(cell.value, str) and cell.value.strip() == "数量":
                    end_row = r - 1
                    break
            if end_row != ws.max_row:
                break

        records_added = 0
        for row_idx in range(start_row, end_row + 1):
            row = ws[row_idx]
            # 姓名列必须有效
            if name_col > len(row):
                continue
            name_cell = row[name_col - 1]  # row是元组，索引0-based
            if not (name_cell.value and DataValidator.is_valid_name(name_cell.value)):
                continue
            name = str(name_cell.value).strip()

            # 处理每个产品块
            for pb in product_blocks:
                q_col = pb['q_col'] - 1
                price_col = pb['price_col'] - 1
                amount_col = pb['amount_col'] - 1
                note_col = pb['note_col'] - 1

                qty = row[q_col].value if q_col < len(row) else None
                price = row[price_col].value if price_col < len(row) else None
                amount = row[amount_col].value if amount_col < len(row) else None
                note = row[note_col].value if note_col < len(row) else ""

                # 判断是否有数据
                has_data = False
                if qty is not None:
                    try:
                        if str(qty).strip() != "" and DataValidator.is_valid_number(qty):
                            has_data = True
                    except:
                        pass
                if not has_data and amount is not None:
                    try:
                        if str(amount).strip() != "" and DataValidator.is_valid_number(amount):
                            has_data = True
                    except:
                        pass
                if not has_data and note is not None:
                    try:
                        if str(note).strip() != "":
                            has_data = True
                    except:
                        pass

                if has_data:
                    # 获取日期：优先从行中日期列读取，否则尝试从表头区域读取日期
                    date_val = None
                    if date_col is not None and date_col <= len(row):
                        date_val = DateParser.parse(row[date_col - 1].value)
                    if not date_val:
                        # 尝试从块上方查找日期（比如A列）
                        for r in range(header_row, max(1, header_row - 10), -1):
                            cell = ws.cell(r, 1)
                            if cell.value:
                                parsed = DateParser.parse(cell.value)
                                if parsed:
                                    date_val = parsed
                                    break
                    if not date_val:
                        # 默认使用今天
                        date_val = datetime.now().strftime("%Y/%m/%d")

                    record = {
                        "日期": date_val,
                        "姓名": name,
                        "批次号": pb['batch'],
                        "产品名称": pb['product'],
                        "数量": float(qty) if qty is not None and DataValidator.is_valid_number(qty) else 0,
                        "计量单位": "",
                        "单价": float(price) if price is not None and DataValidator.is_valid_number(price) else 0,
                        "金额": float(amount) if amount is not None and DataValidator.is_valid_number(amount) else 0,
                        "车间名称": self.sheet_name,
                        "备注": str(note) if note is not None else ""
                    }
                    if record["金额"] == 0 and record["数量"] and record["单价"]:
                        record["金额"] = record["数量"] * record["单价"]
                    if DataValidator.validate_record(record):
                        data_list.append(record)
                        records_added += 1
        return records_added

    def _fallback_extract(self, ws, data_list):
        """子类可重写"""
        pass

    def _extract_with_manual_config(self, ws, data_list, config):
        """手动配置提取（单块模式，不处理多表头）"""
        header_row = config['header_row']
        data_start = config['data_start_row']
        date_col = config.get('date_col')
        name_col = config.get('name_col')
        batch_col = config.get('batch_col')
        product_col = config.get('product_col')
        qty_col = config.get('qty_col')
        price_col = config.get('price_col')
        amount_col = config.get('amount_col')
        note_col = config.get('note_col')

        # 尝试从顶部获取日期
        default_date = None
        for r in range(1, 11):
            for c in range(1, 6):
                cell = ws.cell(r, c)
                if cell.value:
                    parsed = DateParser.parse(cell.value)
                    if parsed:
                        default_date = parsed
                        break
            if default_date:
                break

        for row_idx in range(data_start, ws.max_row + 1):
            row = ws[row_idx]
            if name_col is not None and name_col < len(row):
                name_val = row[name_col].value
                if not DataValidator.is_valid_name(name_val):
                    continue
                name = str(name_val).strip()
            else:
                continue

            batch = "0"
            if batch_col is not None and batch_col < len(row):
                batch = str(row[batch_col].value).strip() if row[batch_col].value else "0"
            product = ""
            if product_col is not None and product_col < len(row):
                product = str(row[product_col].value).strip() if row[product_col].value else ""
            if not product:
                continue

            qty = row[qty_col].value if qty_col is not None and qty_col < len(row) else 0
            price = row[price_col].value if price_col is not None and price_col < len(row) else 0
            amount = row[amount_col].value if amount_col is not None and amount_col < len(row) else 0
            note = row[note_col].value if note_col is not None and note_col < len(row) else ""

            has_data = False
            if qty is not None:
                try:
                    if str(qty).strip() != "" and DataValidator.is_valid_number(qty):
                        has_data = True
                except:
                    pass
            if not has_data and amount is not None:
                try:
                    if str(amount).strip() != "" and DataValidator.is_valid_number(amount):
                        has_data = True
                except:
                    pass
            if not has_data and note:
                has_data = True

            if has_data:
                date_val = None
                if date_col is not None and date_col < len(row):
                    date_val = DateParser.parse(row[date_col].value)
                if not date_val:
                    date_val = default_date
                record = {
                    "日期": date_val,
                    "姓名": name,
                    "批次号": batch,
                    "产品名称": product,
                    "数量": float(qty) if qty is not None and DataValidator.is_valid_number(qty) else 0,
                    "计量单位": "",
                    "单价": float(price) if price is not None and DataValidator.is_valid_number(price) else 0,
                    "金额": float(amount) if amount is not None and DataValidator.is_valid_number(amount) else 0,
                    "车间名称": self.sheet_name,
                    "备注": str(note) if note is not None else ""
                }
                if record["金额"] == 0 and record["数量"] and record["单价"]:
                    record["金额"] = record["数量"] * record["单价"]
                if DataValidator.validate_record(record):
                    data_list.append(record)


# ============================
# 车间专用提取器（可继承，但当前动态提取已通用）
# ============================
class RaorouExtractor(WorkshopDataExtractor):
    pass

class ZhizuoExtractor(WorkshopDataExtractor):
    pass

class BaozhuangExtractor(WorkshopDataExtractor):
    # 包装可保留固定块逻辑，但动态解析也能处理，故留空
    pass


# ============================
# 工具函数：保存结果
# ============================
def save_to_output(data_list):
    if not data_list:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "数据收集表"
    headers = ["日期", "姓名", "批次号", "产品名称", "数量", "计量单位", "单价", "金额", "车间名称", "备注"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    for i, data in enumerate(data_list, 2):
        for col, key in enumerate(headers, 1):
            value = data.get(key, "")
            ws.cell(row=i, column=col, value=value)
    output_buffer = BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


# ============================
# Streamlit 界面
# ============================
def main():
    st.set_page_config(page_title="车间日报提取工具（多表头自适应）", layout="wide")
    st.title("🏭 车间生产日报数据处理系统（多表头自适应）")
    st.markdown("""
    **使用说明：**
    - 上传车间日报表文件（支持 .xlsx, .xls）。
    - 系统会自动识别所有表头块（每天独立表头），并提取数据。
    - 若自动识别不理想，可勾选 **“手动配置字段映射”** 进行列指定。
    - 支持多文件批量处理，结果汇总下载。
    """)
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "📤 请选择要处理的文件 (可多选)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )

    # 如果上传了文件，加载预览数据（第一个文件的第一个工作表）
    if uploaded_files and 'preview_df' not in st.session_state:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(uploaded_files[0].getbuffer())
                tmp_path = tmp.name
            wb = load_workbook(tmp_path, data_only=True)
            sheet = wb.active
            data_rows = []
            for row in sheet.iter_rows(max_row=20, values_only=True):
                data_rows.append(row)
            df = pd.DataFrame(data_rows)
            st.session_state.preview_df = df
            wb.close()
            os.unlink(tmp_path)
        except Exception as e:
            st.error(f"预览加载失败: {e}")

    # 自动处理按钮
    if st.button("🚀 开始自动处理", type="primary"):
        if not uploaded_files:
            st.warning("⚠️ 请先上传至少一个文件！")
        else:
            all_data = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            total_files = len(uploaded_files)

            for idx, uploaded_file in enumerate(uploaded_files):
                progress_bar.progress((idx + 1) / total_files)
                status_text.text(f"正在处理: {uploaded_file.name} ...")
                if "车间" not in uploaded_file.name and "生产日报" not in uploaded_file.name:
                    st.info(f"⏭️ 文件 '{uploaded_file.name}' 不包含关键字，已跳过。")
                    continue
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                        tmp.write(uploaded_file.getbuffer())
                        tmp_path = tmp.name
                    wb = load_workbook(tmp_path, data_only=True)
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        # 根据车间名选择提取器（但动态提取通用，可不区分）
                        if "绕肉" in sheet_name:
                            extractor = RaorouExtractor(sheet_name)
                        elif "制作" in sheet_name:
                            extractor = ZhizuoExtractor(sheet_name)
                        elif "包装" in sheet_name or "挑选" in sheet_name:
                            extractor = BaozhuangExtractor(sheet_name)
                        else:
                            extractor = WorkshopDataExtractor(sheet_name)
                        extractor.extract(ws, all_data)
                    wb.close()
                    os.unlink(tmp_path)
                except Exception as e:
                    st.error(f"❌ 处理文件 {uploaded_file.name} 时发生错误: {str(e)}")

            if all_data:
                st.success(f"✅ 自动处理完成！共提取 **{len(all_data)}** 条记录。")
                st.session_state.all_data = all_data
                output_buffer = save_to_output(all_data)
                st.download_button(
                    label="📥 下载结果文件 (生产车间统计数据收集.xlsx)",
                    data=output_buffer,
                    file_name="生产车间统计数据收集.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ 自动提取未获得任何数据，请尝试手动配置。")

    # 手动配置区域（始终显示）
    st.markdown("---")
    with st.expander("🔧 手动配置字段映射（高级）", expanded=False):
        st.info("若自动提取不理想，请根据下方预览数据指定各字段的列索引（从0开始）。")
        if 'preview_df' in st.session_state:
            st.dataframe(st.session_state.preview_df)
            col1, col2 = st.columns(2)
            with col1:
                header_row = st.number_input("表头所在行（0-based）", min_value=0, value=0, step=1)
                data_start_row = st.number_input("数据起始行（0-based）", min_value=0, value=1, step=1)
            with col2:
                date_col = st.number_input("日期列索引（-1表示无）", min_value=-1, value=-1, step=1)
                name_col = st.number_input("姓名列索引", min_value=0, value=1, step=1)
                batch_col = st.number_input("批次号列索引（-1表示无）", min_value=-1, value=-1, step=1)
                product_col = st.number_input("产品名称列索引", min_value=0, value=2, step=1)
                qty_col = st.number_input("数量列索引", min_value=0, value=3, step=1)
                price_col = st.number_input("单价列索引", min_value=0, value=4, step=1)
                amount_col = st.number_input("金额列索引", min_value=0, value=5, step=1)
                note_col = st.number_input("备注列索引（-1表示无）", min_value=-1, value=-1, step=1)

            if st.button("🔄 使用手动配置重新提取"):
                config = {
                    'header_row': header_row,
                    'data_start_row': data_start_row,
                    'date_col': date_col if date_col >= 0 else None,
                    'name_col': name_col,
                    'batch_col': batch_col if batch_col >= 0 else None,
                    'product_col': product_col,
                    'qty_col': qty_col,
                    'price_col': price_col,
                    'amount_col': amount_col,
                    'note_col': note_col if note_col >= 0 else None
                }
                if not uploaded_files:
                    st.warning("请先上传文件。")
                else:
                    all_data_manual = []
                    try:
                        # 处理第一个文件（可扩展为选择文件）
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                            tmp.write(uploaded_files[0].getbuffer())
                            tmp_path = tmp.name
                        wb = load_workbook(tmp_path, data_only=True)
                        for sheet_name in wb.sheetnames:
                            ws = wb[sheet_name]
                            extractor = WorkshopDataExtractor(sheet_name)
                            extractor.extract(ws, all_data_manual, manual_config=config)
                        wb.close()
                        os.unlink(tmp_path)
                    except Exception as e:
                        st.error(f"手动提取出错: {e}")
                    if all_data_manual:
                        st.success(f"手动提取成功，共 {len(all_data_manual)} 条记录。")
                        st.session_state.all_data = all_data_manual
                        output_buffer = save_to_output(all_data_manual)
                        st.download_button(
                            label="📥 下载手动提取结果",
                            data=output_buffer,
                            file_name="手动提取结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("手动提取未获得数据，请检查配置。")
        else:
            st.info("请先上传文件以加载预览数据。")

if __name__ == "__main__":
    main()