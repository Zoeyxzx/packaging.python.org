Discussionsimport datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 创建工作簿和工作表
wb = Workbook()
ws = wb.active
ws.title = "学习计划"

# 定义列标题（周一到周日）
headers = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
ws.append(headers)

# 设置列标题样式：粗体、居中
header_font = Font(bold=True)
header_alignment = Alignment(horizontal="center", vertical="center")
for col_num, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.alignment = header_alignment

# 应用过滤器
ws.auto_filter.ref = "A1:G1"

# 日期范围：从三月十二号到四月十三号（此处以 2025 年为例）
start_date = datetime.date(2025, 3, 12)
end_date = datetime.date(2025, 4, 13)
delta = datetime.timedelta(days=1)

# 按周存放数据，每周为一行
week_rows = []
current_week = [""] * 7  # Monday: index 0 ... Sunday: index 6
current_date = start_date

while current_date <= end_date:
    weekday = current_date.weekday()  # Monday=0, Tuesday=1, ..., Sunday=6
    # 构造单元格内容：日期字符串
    cell_text = current_date.strftime("%Y-%m-%d")
    
    # 添加课程安排：
    # 每周三下午 2:00-4:00 有数学课
    if weekday == 2:
        cell_text += "\nMath 14:00-16:00"
    # 每周六和周天早上 9:00-11:00 有物理课
    if weekday in [5, 6]:
        cell_text += "\nPhysics 09:00-11:00"
    
    current_week[weekday] = cell_text
    
    # 遇到周日或到达最后一天则将该周加入数据
    if weekday == 6:
        week_rows.append(current_week)
        current_week = [""] * 7
    current_date += delta

# 如果最后一周不满七天，也加入
if any(entry != "" for entry in current_week):
    week_rows.append(current_week)

# 将每周数据写入工作表，从第二行开始
for week in week_rows:
    ws.append(week)

# 设置所有单元格的边框（细边框）
thin = Side(style="thin")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=7):
    for cell in row:
        cell.border = border
        # 数据对齐：文本左对齐（这里数据均为文本）
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 自动调整列宽
for col in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value:
            cell_length = len(str(cell.value))
            if cell_length > max_length:
                max_length = cell_length
    ws.column_dimensions[col_letter].width = max_length + 2

# 保存为 .xlsx 文件
file_name = "Zoey三月四月学习计划.xlsx"
wb.save(file_name)
print(f"文件已保存为 {file_name}")

###########

**Discussions** are focused on providing comprehensive information about a
specific topic. If you're just trying to get stuff done, see
:doc:`/guides/index`.

.. toctree::
   :maxdepth: 1

   versioning
   deploying-python-applications
   pip-vs-easy-install
   install-requires-vs-requirements
   distribution-package-vs-import-package
   package-formats
   src-layout-vs-flat-layout
   setup-py-deprecated
   single-source-version
   downstream-packaging
