# 周计划工作台

导入供电所「表1」Excel，生成：

- 周计划明细 Excel（A3 横向、不跨行分页、管控措施行高可调）
- 三四五级风险汇总 TXT
- 周计划项目汇总 TXT

## 怎么下载

1. 打开本仓库：https://github.com/maxinxi/weekly-plan-workbench
2. 点绿色 **Code** → **Download ZIP**
3. 解压即可用

## 用法一：HTML 工作台（推荐，不用装 Python）

解压后双击 **周计划工作台.html**。

按钮：

| 按钮 | 作用 |
|---|---|
| 载入示例 | 恢复内置示例周 |
| 点击或拖入 表1.xlsx | 导入源表 |
| 下载示例表1 | 下载一份可导入的示例源表 |
| 管控措施行高滑条 | 改 +1 / +2 / +0 |
| 下载明细 Excel | 导出周计划明细 |
| 三四五级风险 TXT | 导出风险汇总 |
| 项目汇总 TXT | 导出项目统计 |

## 用法二：Python 脚本（批量处理目录里的表1）

```bash
pip install -r requirements.txt
python weekly-plan-export.py
```

把表1.xlsx 和脚本放在同一目录再运行。

改管控措施行高只改脚本最上面这一行：

```python
MEASURES_HEIGHT_PLUS = 1.0   # 2.0 = +2，0 = 不加
```

原脚本「全表统一 -1」已取消，避免裁切管控措施内容。
