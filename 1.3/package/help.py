import tkinter as tk
from tkinter import Scrollbar, Text
import os
import tkinter.font as tkFont


class SyntenyGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("帮助说明")
        self.master.geometry("800x600")

        # 设置窗口图标
        icon_file = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_file):
            try:
                self.master.iconbitmap(icon_file)
            except Exception as e:
                print(f"加载图标失败: {e}")

        self.create_help_content()

    def create_help_content(self):
        help_window = self.master
        # 滚动条 + 文本框
        scrollbar = Scrollbar(help_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 设置宋体五号字体 (约12pt)
        font_style = tkFont.Font(family="SimSun", size=14)  # SimSun=宋体

        text = Text(help_window, wrap="word", yscrollcommand=scrollbar.set, font=font_style)
        text.pack(expand=True, fill="both")
        scrollbar.config(command=text.yview)

        # 使用说明内容
        help_content = """
基因工具 — 详细使用说明书

版本说明
---------
版本：1.3
更新说明：
- 增强了图形界面交互体验
- 新增基因关联/网络可视化模块
- 优化竖向匹配逻辑并增加导出颜色标记
- 其它若干稳定性与兼容性修复

目录
---------
1. 概述
2. 环境与依赖
3. 快速开始
4. 各功能模块详细使用
   4.1 文件预处理（基因ID / 信息表格）
   4.2 基因匹配（普通版）
   4.3 基因匹配 Pro（增强版）
   4.4 共线性/ID 文件转换
   4.5 关联/网络可视化
   4.6 共线性可视化（热图 / 散点图）
5. 输出文件说明
6. 大数据与性能优化建议
7. 常见问题（FAQ）
8. 编码、字体与跨平台注意项
9. 联系与意见反馈

1. 概述
---------
本工具集面向基因表格的比对、匹配与可视化工作，支持 Excel/CSV 数据格式。主要用于：
- 在填表数据与信息表之间执行精确/模糊匹配
- 导出匹配结果（Excel，含颜色高亮）
- 解析 collinearity/ID 风格文本并转为 Excel
- 可视化基因共线性与基因关联网络

2. 环境与依赖
---------
建议 Python 版本：3.7+
必需库（可通过 pip 安装）：
- pandas
- numpy
- matplotlib
- networkx
- openpyxl
- xlsxwriter（用于带颜色的 Excel 输出）
说明：若使用打包版本（PyInstaller），请在打包时将依赖一并包含。

安装示例：
pip install pandas numpy matplotlib networkx openpyxl xlsxwriter

3. 快速开始
---------
1）启动程序：
   python main.py
   或直接运行打包后的可执行文件。
2）在主窗口菜单选择所需功能：
   - 文件预处理 -> 基因ID / 信息表格
   - 工具 -> 基因匹配 / 基因匹配pro
   - 可视化 -> 共线性可视化 / 关系网络图（基因关联）
3）根据界面提示选择输入文件、设置列后运行生成并保存结果。

4. 各功能模块详细使用
---------

4.1 文件预处理（基因ID / 信息表格）
- 目的：将原始表格转换为标准化格式，便于后续匹配。
- 输入：Excel（.xlsx/.xls）或 CSV。
- 操作：
  1. 打开“文件预处理”菜单，选择“信息表格”或“基因ID”。
  2. 浏览并选择待处理文件，指定输出目录。
  3. 点击“开始”或“转换”按钮，等待完成提示。
- 输出：标准化后的 Excel 文件，保存在指定目录。

4.2 基因匹配（普通版）
- 目的：在“填入表格”（A）与“信息表格”（B）之间进行匹配。
- 功能点：
  - 精确匹配：A 中值 == B 中值
  - 简单模糊匹配：包含关系（字符串包含/被包含）
  - 导出 B 表中除匹配列外的其他列（可选）
- 操作：
  1. 选择 A、B 文件与输出文件夹。
  2. 选择 A 表目标列、B 表匹配列（及可选信息列）。
  3. 点击“开始匹配”。
- 输出：Excel 文件，含“匹配形式”（完全匹配/模糊匹配/未匹配）、匹配结果字段与可选附带的 B 表其他列。

4.3 基因匹配 Pro（增强版）
- 说明：在普通版基础上优化了性能，支持多匹配复制 A 行并生成多行记录；并对差异情况（A 多 / A 少 / 长度相同）做着色保存到 Excel（需 xlsxwriter）。
- 操作与普通版相似，但提供更多选项（字符差异容忍度、是否导出 B 表其他列等）。
- 输出：带颜色标注的 Excel（不同差异使用不同背景色），便于人工检查。

4.4 共线性/ID 文件转换
- 功能：解析 .collinearity 或自定义格式的文本（通常为基因比对/共线性输出），转换为 Excel。
- 操作：
  1. 打开“ID表格转换”或共线性解析界面。
  2. 选择 .txt/.collinearity 文件与输出目录。
  3. 开始转换，完成后会显示生成的 Excel 路径并提示是否打开。
- 注意：解析基于常见行格式（含“## Alignment <num>”块标识、注释以 # 开头、以及常见 2/3 列数据行）。若格式差异较大，请先预处理或手工调整文本。

4.5 关联/网络可视化（基因关联）
- 目标：将节点对（如 geneA、geneB）绘制为关系网络，可选择权重列用于加权边。
- 操作：
  1. 选择包含节点对的表格文件（Excel/CSV）。
  2. 在“节点1列/节点2列”下拉中选择对应列；如有权重，选择权重列。
  3. 点击“生成图形”以预览网络；可保存为 PNG。
- 参数与提示：
  - 若选择权重列，程序会按 (node1,node2) 聚合并取权重均值作为边权值。
  - 对于节点数非常多的数据，布局计算（spring_layout）可能较慢，请考虑先筛选子网络或限制绘制节点数。
  - 图形颜色映射表示节点度（连接数），并可显示边权数值（两位小数）。

4.6 共线性可视化（热图 / 散点图）
- 热图：基于 X、Y、数据列构建 pivot_table（aggfunc=mean），并用 imshow 绘制色块图。
- 散点图：选择数值列作为 X、Y 绘制散点，支持设置轴标签与标题。
- 注意：
  - 热图依赖于数据的透视表结构，若类别过多可能导致显示拥挤或内存问题。
  - 散点图会把非数值转换为 NaN 并剔除，确保选择的列包含数值。

5. 输出文件说明
---------
- Excel 输出：命名规则通常为 inputname_功能_时间戳.xlsx 或 output_optimized_match_YYYYMMDD_HHMMSS.xlsx。
- Pro 版本会生成带颜色的 Excel（需要 xlsxwriter），颜色表示差异类型（A多/蓝、A少/红、长度相同/黄）。
- PNG 输出：保存绘图为高分辨率 PNG（dpi=300），文件名含输入文件名与图类型。

6. 大数据与性能优化建议
---------
- 对于行数 > 50k 的表格：
  - 在匹配前尽量筛选或去重，减少无效比较。
  - 使用 CSV 而非 Excel 有时更快；读取 CSV 时指定合适的 encoding 与 dtype。
  - 避免在 UI 线程进行大量计算，程序已对关键流程使用后台线程（Pro 版本亦如此）。
- 可视化：
  - networkx spring_layout 对大图耗时显著，可考虑降采样节点或只绘制核心子网。

7. 常见问题（FAQ）
---------
Q：程序卡顿或无响应？
A：可能是处理大型数据或绘图计算。建议等待任务结束，或在小样本上测试参数。若卡住可终止任务并在日志/控制台查看异常信息。

Q：Excel 保存失败或颜色不显示？
A：请确保已安装 xlsxwriter。若颜色不生效，可能使用的 Excel 引擎为 openpyxl，Pro 版本需要 xlsxwriter 才能写入单元格格式。

Q：中文显示乱码或字体问题？
A：Matplotlib 默认字体可能不支持中文，程序尝试设置 SimHei（宋体）。若系统无 SimHei，请安装中文字体或修改程序中 plt.rcParams 的字体配置。

Q：如何改变热图的聚合方法？
A：当前为 mean，如需 median 或 sum，可手动在代码 pivot_table 的 aggfunc 参数中修改为 'median' 或 'sum'。

8. 编码、字体与跨平台注意项
---------
- CSV 编码：Windows 常见为 GBK（或 CP936）；Linux/Unix 常见为 UTF-8。若读取失败，请尝试手工指定编码或另存为 UTF-8。
- 字体：若遇中文标签显示为空或问号，需在系统安装 SimHei 等中文字体，或在代码中设置其它可用中文字体。
- 打包：使用 PyInstaller 打包时，将依赖库与字体文件一并包含，避免运行环境缺少引擎或字体。

9. 联系与意见反馈
---------
如需功能扩展、错误修复或定制化服务，请将以下信息一并提供：
- 问题重现步骤与最小示例数据（可脱敏）
- Python 版本与操作系统（Windows/macOS/Linux）
- 报错堆栈或程序运行日志
开发者邮箱/联系方式：请在项目 README 或版本信息中查看（或通过项目维护渠道联系）。

附：良好使用习惯
---------
1. 在处理前备份原始数据文件，避免误操作导致数据丢失。
2. 对于大规模数据，先在一小部分样本上调参测试结果。
3. 使用明确的目录结构保存输入/输出，便于版本管理与结果追踪。

感谢使用本工具，欢迎反馈使用体验与建议。
"""
        text.insert("1.0", help_content)
        text.config(state="disabled")  # 只读