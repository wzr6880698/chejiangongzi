import openpyxl
from openpyxl import load_workbook
import os
from datetime import datetime, timedelta
import re
import streamlit as st
from io import BytesIO
import tempfile
from collections import defaultdict


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
# 核心提取器
# ============================
class WorkshopDataExtractor:
    def __init__(self, sheet_name):
        self.sheet_name = sheet_name

    def extract(self, ws, data_list):
        self._try_dynamic_extract(ws, data_list)

    def _try_dynamic_extract(self, ws, data_list):
        header_blocks = self._find_header_blocks(ws)
        if not header_blocks:
            return False
        total = 0
        for block in header_blocks:
            total += self._extract_block_data(ws, block, data_list)
        return total > 0

    def _get_value_from_label(self, ws, label_row, label_col, max_row, max_col):
        """
        根据关键词所在的单元格，智能获取其对应的值。
        策略：
        1. 检查右侧单元格，若右侧非空且不是另一个标签，则取右侧。
        2. 否则，检查下方单元格，取下方。
        """
        label_keywords = {
            "批次号", "批号", "产品名称", "品名", "日期", "生产日期",
            "姓名", "数量", "单价", "金额", "备注", "序号", "合计"
        }

        # 检查右侧
        right_cell = ws.cell(label_row, label_col + 1) if label_col + 1 <= max_col else None
        if right_cell and right_cell.value is not None:
            right_val = right_cell.value
            if isinstance(right_val, str):
                right_val = right_val.strip()
            if right_val and right_val not in label_keywords:
                return str(right_val)

        # 检查下方
        down_cell = ws.cell(label_row + 1, label_col) if label_row + 1 <= max_row else None
        if down_cell and down_cell.value is not None:
            down_val = down_cell.value
            if isinstance(down_val, str):
                down_val = down_val.strip()
            if down_val and str(down_val).strip() != "":
                return str(down_val)

        return None

    def _find_header_blocks(self, ws):
        blocks = []
        max_row = ws.max_row
        max_col = ws.max_column

        quantity_cells = []
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                if cell.value and isinstance(cell.value, str) and cell.value.strip() == "数量":
                    quantity_cells.append((r, c))

        if not quantity_cells:
            return []

        rows_with_q = defaultdict(list)
        for r, c in quantity_cells:
            rows_with_q[r].append(c)

        for header_row, q_cols in sorted(rows_with_q.items()):
            # 查找姓名列
            name_col = None
            for r in range(header_row, max(1, header_row - 5), -1):
                for c in range(1, max_col + 1):
                    cell = ws.cell(r, c)
                    if cell.value and isinstance(cell.value, str) and cell.value.strip() == "姓名":
                        name_col = c
                        break
                if name_col:
                    break
            if name_col is None:
                name_col = 2

            # 公共元数据
            search_radius_row = 3
            search_radius_col = 5
            block_metadata = {'date': None, 'batch': None, 'product': None}

            # 日期
            for r in range(header_row - search_radius_row, header_row + search_radius_row + 1):
                if r < 1 or r > max_row:
                    continue
                for c in range(1, max_col + 1):
                    cell = ws.cell(r, c)
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value.strip()
                        if val in ["日期", "生产日期"]:
                            value_str = self._get_value_from_label(ws, r, c, max_row, max_col)
                            if value_str:
                                parsed = DateParser.parse(value_str)
                                if parsed:
                                    block_metadata['date'] = parsed
                                    break
                    if block_metadata['date']:
                        break
                if block_metadata['date']:
                    break

            if not block_metadata['date']:
                for r in range(header_row - search_radius_row, header_row + search_radius_row + 1):
                    if r < 1 or r > max_row:
                        continue
                    for c in range(1, 6):
                        cell = ws.cell(r, c)
                        if cell.value:
                            parsed = DateParser.parse(cell.value)
                            if parsed:
                                block_metadata['date'] = parsed
                                break
                    if block_metadata['date']:
                        break

            # 批次号
            for r in range(header_row - search_radius_row, header_row + search_radius_row + 1):
                if r < 1 or r > max_row:
                    continue
                for c in range(1, max_col + 1):
                    cell = ws.cell(r, c)
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value.strip()
                        if val in ["批次号", "批号"]:
                            value_str = self._get_value_from_label(ws, r, c, max_row, max_col)
                            if value_str:
                                block_metadata['batch'] = value_str
                                break
                    if block_metadata['batch']:
                        break
                if block_metadata['batch']:
                    break

            # 产品名称
            for r in range(header_row - search_radius_row, header_row + search_radius_row + 1):
                if r < 1 or r > max_row:
                    continue
                for c in range(1, max_col + 1):
                    cell = ws.cell(r, c)
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value.strip()
                        if val in ["产品名称", "品名"]:
                            value_str = self._get_value_from_label(ws, r, c, max_row, max_col)
                            if value_str:
                                block_metadata['product'] = value_str
                                break
                    if block_metadata['product']:
                        break
                if block_metadata['product']:
                    break

            # 处理每个数量列
            product_blocks = []
            for q_col in sorted(q_cols):
                batch = block_metadata['batch'] if block_metadata['batch'] is not None else "0"
                product = block_metadata['product'] if block_metadata['product'] is not None else f"产品{q_col}"
                date_val = block_metadata['date']

                if block_metadata['batch'] is None:
                    for r in range(header_row - 2, header_row + 3):
                        if r < 1 or r > max_row:
                            continue
                        for c in range(max(1, q_col - 2), min(max_col, q_col + 3) + 1):
                            cell = ws.cell(r, c)
                            if cell.value and isinstance(cell.value, str):
                                val = cell.value.strip()
                                if val in ["批次号", "批号"]:
                                    value_str = self._get_value_from_label(ws, r, c, max_row, max_col)
                                    if value_str:
                                        batch = value_str
                                        break
                        if batch != "0":
                            break

                if block_metadata['product'] is None:
                    for r in range(header_row - 2, header_row + 3):
                        if r < 1 or r > max_row:
                            continue
                        for c in range(max(1, q_col - 2), min(max_col, q_col + 3) + 1):
                            cell = ws.cell(r, c)
                            if cell.value and isinstance(cell.value, str):
                                val = cell.value.strip()
                                if val in ["产品名称", "品名"]:
                                    value_str = self._get_value_from_label(ws, r, c, max_row, max_col)
                                    if value_str:
                                        product = value_str
                                        break
                        if product != f"产品{q_col}":
                            break

                price_col = q_col + 1
                amount_col = q_col + 2
                note_col = q_col + 3
                product_blocks.append({
                    'q_col': q_col,
                    'price_col': price_col,
                    'amount_col': amount_col,
                    'note_col': note_col,
                    'batch': batch,
                    'product': product,
                    'date': date_val
                })

            blocks.append({
                'header_row': header_row,
                'name_col': name_col,
                'product_blocks': product_blocks
            })

        return blocks

    def _extract_block_data(self, ws, block, data_list):
        header_row = block['header_row']
        name_col = block['name_col']
        product_blocks = block['product_blocks']

        start_row = header_row + 1
        end_row = ws.max_row
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
            if name_col > len(row):
                continue
            name_cell = row[name_col - 1]
            if not (name_cell.value and DataValidator.is_valid_name(name_cell.value)):
                continue
            name = str(name_cell.value).strip()

            for pb in product_blocks:
                q_col = pb['q_col'] - 1
                price_col = pb['price_col'] - 1
                amount_col = pb['amount_col'] - 1
                note_col = pb['note_col'] - 1

                qty = row[q_col].value if q_col < len(row) else None
                price = row[price_col].value if price_col < len(row) else None
                amount = row[amount_col].value if amount_col < len(row) else None
                note = row[note_col].value if note_col < len(row) else ""

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
                    date_val = pb.get('date')
                    if not date_val:
                        for c in range(1, 6):
                            cell = ws.cell(row_idx, c)
                            if cell.value:
                                parsed = DateParser.parse(cell.value)
                                if parsed:
                                    date_val = parsed
                                    break
                    if not date_val:
                        for r in range(header_row, max(1, header_row - 10), -1):
                            cell = ws.cell(r, 1)
                            if cell.value:
                                parsed = DateParser.parse(cell.value)
                                if parsed:
                                    date_val = parsed
                                    break
                    if not date_val:
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


# ============================
# 车间专用提取器
# ============================
class RaorouExtractor(WorkshopDataExtractor):
    pass


class ZhizuoExtractor(WorkshopDataExtractor):
    pass


class BaozhuangExtractor(WorkshopDataExtractor):
    pass


# ============================
# 保存结果
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
    st.set_page_config(page_title="车间日报提取工具（智能值定位）", layout="wide")
    st.title("🏭 车间生产日报数据处理系统（智能值定位）")
    st.markdown("""
    **使用说明：**
    - 上传车间日报表文件（支持 .xlsx, .xls）。
    - 系统自动识别数据块，寻找关键词后：
        - 若右侧单元格不是标签，则取右侧值（二行/三行表头）；
        - 否则取下方单元格值（一行表头）。
    - 兼容一行、两行、三行及任意行表头。
    - 支持多文件批量处理，结果汇总下载。
    """)
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "📤 请选择要处理的文件 (可多选)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )

    if st.button("🚀 开始处理", type="primary"):
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
                    with tempfile.NamedTemporaryFile(delete=False,
                                                     suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                        tmp.write(uploaded_file.getbuffer())
                        tmp_path = tmp.name
                    wb = load_workbook(tmp_path, data_only=True)
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
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
                st.success(f"✅ 处理完成！共提取 **{len(all_data)}** 条记录。")
                output_buffer = save_to_output(all_data)
                st.download_button(
                    label="📥 下载结果文件 (生产车间统计数据收集.xlsx)",
                    data=output_buffer,
                    file_name="生产车间统计数据收集.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ 未能提取到任何数据，请检查文件格式。")


if __name__ == "__main__":
    main()