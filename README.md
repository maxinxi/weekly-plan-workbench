# 周计划工作台

**只下 HTML 即可使用。无需安装 Python，无需联网。**

[下载单文件「周计划工作台.html」](https://raw.githubusercontent.com/maxinxi/weekly-plan-workbench/main/%E5%91%A8%E8%AE%A1%E5%88%92%E5%B7%A5%E4%BD%9C%E5%8F%B0.html)：右键链接另存为，或打开仓库中的同名文件，选择 Download raw file。保存后用 Edge / Chrome 打开。

## 使用

1. 拖入原始「表1.xlsx」。网页读取文件，在本机浏览器处理。
2. 查看处理日志；能确定的换行、时间分隔符自动修复。日期、时间等不能确定的问题保留原值并报告。
3. 点击 **下载处理后源表**，得到 A3 横向的完整源表；管控措施保留，超高行自动插行。
4. 点击 **下载明细 Excel**，得到周计划明细。**周计划明细不带管控措施。** 风险汇总和项目汇总仍可单独下载。
5. 点击 **源文件备份 ZIP** 保存原件备份；需要核对时下载源表处理报告。

默认采用第37周成品的排版方向：保留已有宽列、原字号和加粗；48磅不会改小。原打印缩放能容纳A3宽度时保留（包括10%），否则设为一页宽、高度不限。保留已有有效打印列范围，每页重复标题和表头，尽量在完整作业之间分页；单个作业过长时在管控措施块之间分页。单块超过一页时明确报告。

「管控措施行高增量」默认 **+1磅**，可改为0或2。取消「保留成品宽列」可使用早期的管控措施列上限56模式；48磅宽表通常应保留默认选项。新增行和安全余量可能增加打印页数，不为凑页数裁掉正文。打印时选择支持A3的打印机或PDF打印机。

## 原件保护与处理顺序

| 顺序 | 行为 |
|---|---|
| 1 | 只读原件，备份一次，建立工作副本 |
| 2 | 保持一行一条预处理，先不插行 |
| 3 | 先逐列估算宽度：中文/全角2.2、ASCII 1.0、余量2.4、下限8；普通自动宽上限48，已足够宽的不缩窄 |
| 4 | 再逐行计算高度：所有单元格需求取最大，横向合并使用列宽总和；纵向合并跳过非首格，利用已有各行高度，仅将不足部分补到首行 |
| 5 | 高度仍超过409.5磅才整行插入；先插行，再复制样式/边框/填充，最后重建纵向合并并扩展原横向合并块 |
| 6 | 输出「（第N周）（处理后的源表）.xlsx」；无周次时用「原名（处理后的源表）.xlsx」 |
| 7 | 成功或失败都清理本次work副本 |

高度按「行数 × 字号 × 1.5 + 6」估算，文字宽度结合字号换算；空行15磅，有内容至少18磅，管控措施另加指定增量。自动换行、垂直居中、显式行高。原表存在远端成万空合并时，只清理工作副本中数据列以外的空白格式，避免虚假大范围拖垮导出。

预处理删除A列为「例/示例」、带删除线的示例行，以及「填报人/填报人及联系方式」行；统一换行并压缩连续空行，人员列姓名与电话跨行时直接拼接。

**标黄不等于示例：三级/3级风险保留；已识别的其他级别（如四级、五级）标黄行删除；无法判断级别时保留并警告。** 删除行涉及已有纵向合并首格时，将原值和原样式迁移到剩余合并区域首格并记录位置。

浏览器受文件权限限制，不会在硬盘创建工作文件夹：首次备份保存在当前浏览器的 IndexedDB，同名备份不覆盖；work是独立内存副本，用后清理。「源文件备份 ZIP」内包含 `源文件备份/原文件名.xlsx`，内容与首次原件逐字节一致。浏览器不允许持久存储时明确警告，改为本次会话备份，请及时下载ZIP。清理浏览器数据或更换浏览器/HTML位置可能使旧备份不可见，因此建议保存ZIP。

## 错误报告

报告包含文件名、原始单元格位置、错误代码和说明：日期与标题/文件名不一致、标题与文件名周次不一致、日期非法、时间缺日期、时刻无法解析、起止顺序异常、必需列/管控措施列缺失、异常合并、输出失败等。不按业务周次擅自推算ISO日期，不猜跨夜日期，不补造时刻。

有数据错误时仍可下载源表排版结果和报告，但暂停生成明细，修正原件后重新导入。异常合并或损坏工作簿会停止源表导出。含公式、表格或条件格式的源表会提示核对引用；不承诺自动改写复杂公式或保留ExcelJS不支持的Excel扩展功能。

## Python参考/批处理版

仅使用HTML无需本节。Python源表处理使用openpyxl；`source_processor.py`是源表入口，`weekly-plan-export.py`保留明细/汇总菜单。

```sh
pip install -r requirements.txt
python source_processor.py "表1（第37周）.xlsx"
python source_processor.py --directory "源表目录" --measures-plus 2
python source_processor.py "表1.xlsx" --strict-widths
python weekly-plan-export.py
```

Python在原文件旁独占创建一次 `源文件备份/原名.xlsx`（存在即跳过），复制到 `work/source-随机名/`，绝不保存回原件；只删除本次工作副本。扫描排除 `~$`、明细、处理后源表、work、源文件备份。运行结束将处理报告写到输出旁边。

## 开发与验证

HTML已内嵌ExcelJS及其JSZip依赖。`source-processor.js`是浏览器源表逻辑，修改后运行构建脚本重新内嵌，交付仍只有一个HTML。构建脚本还修复了内置ExcelJS的sheetPr节点顺序，防止桌面Excel拒绝打开源表。

```sh
python tools/build_html.py
python -m unittest discover -s tests -v
node tests/browser_test.cjs
node tests/browser_test.cjs "测试表1.xlsx" "work/browser-test"
python tests/verify_browser_output.py "测试表1.xlsx" "work/browser-test/browser.xlsx"
```

覆盖一次备份、异常清理、删行与三级风险例外、列宽→行高→插行顺序、每行显式高度、409.5上限、横/纵合并保护、48磅加粗保留、日期时间报错、命名、空白格式膨胀、明细不含管控措施。浏览器产物由openpyxl独立回读，与Python结果比较。可选安装Playwright后运行 `node tests/offline_ui.cjs`，用本机Edge验证断网导入、下载和跨刷新备份。

真实第36/37周文件在本机核验，未加入仓库。第37周样本保留所有记录及三级风险黄色标记，默认增加7行；用桌面Excel导出为6页A3，以留出正文空间。用户提供的5页PDF用于版式对照。

参考：[openpyxl插删行与合并](https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html)、[ExcelJS合并单元格及打印设置](https://github.com/exceljs/exceljs)、[Microsoft Open XML SheetProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetproperties)。实现采用“保护合并关系后移动行”“复制样式后重建合并”“按内容与打印块分页”三条原则。
