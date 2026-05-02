# 自动化入账脚本

本仓库保存两份基于 SAP GUI Scripting 和 Excel 的自动化入账脚本：

- `scripts/结汇入账.py`：SAP F-65 银行外币入账自动化。
- `scripts/F-44.py`：SAP F-44 供应商清账自动化。

## 运行环境

仅支持 Windows 桌面环境，并需要本机已经安装和配置：

- Python 3.9 或以上版本。
- Microsoft Excel。
- SAP GUI，并启用 SAP GUI Scripting。
- 已登录 SAP，且当前窗口有一个可操作的会话。

Python 依赖：

```powershell
pip install -r requirements.txt
```

## SAP GUI 准备

运行脚本前请先确认：

1. 已登录正确 SAP 系统和公司代码对应环境。
2. SAP GUI Scripting 已启用。
3. 当前 SAP 窗口没有未处理的弹窗。
4. Excel 文件没有被其他程序锁定。
5. 建议先用少量测试数据执行，确认字段、期间、币种、科目和权限都正确。

## F-65 结汇入账

脚本位置：

```powershell
python .\scripts\结汇入账.py
```

运行前需要在脚本底部修改 Excel 路径：

```python
EXCEL_PATH = r"C:\Users\3006699\PycharmProjects\PythonProject\自动化入账\F65_Data.xlsx"
SHEET_NAME = None
```

`SHEET_NAME = None` 表示读取第一个工作表；如需指定工作表，改为工作表名称字符串。

Excel 必填表头：

| 字段 | 说明 |
| --- | --- |
| 凭证日期 | SAP 日期，脚本会尝试格式化为 `dd.mm.yyyy` |
| 公司代码 | SAP 公司代码 |
| 货币 | 外币币种 |
| 借方科目 | 借方银行/科目 |
| 借方金额 | 借方外币金额 |
| 借方本币金额 | 借方本币金额 |
| 贷方科目 | 贷方银行/科目 |
| 贷方金额 | 贷方外币金额 |
| 贷方本币金额 | 贷方本币金额 |

脚本会校验借方金额和贷方金额是否一致。处理结果会回写到 Excel：

- `处理结果`
- `SAP凭证号`
- `错误信息`
- `处理时间`

日志默认写入：

```text
C:\SAP_Automation\Logs
```

## F-44 供应商清账

脚本位置：

```powershell
python .\scripts\F-44.py
```

运行前需要在脚本底部修改 Excel 路径：

```python
excel_file = r"C:\Users\3006699\Desktop\供应商清账数据.xlsx"
runner = SAPF44Auto(excel_file)
```

Excel 从第 2 行开始读取数据，列定义如下：

| 列 | 字段 | 默认值 |
| --- | --- | --- |
| A | 供应商代码 | 必填 |
| B | 过账日期 `BUDAT` | 当天日期 |
| C | 期间 `MONAT` | 当前月份 |
| D | 公司代码 `BUKRS` | `3401` |
| E | 货币 `WAERS` | `RMB` |
| F | 特别总账标识 `AGUMS` | `LKJM` |
| G | 抬头文本 | `清账` |

脚本会回写处理结果：

| 列 | 字段 |
| --- | --- |
| H | `result` |
| I | `message` |
| J | `processed_at` |

## 常见问题

- 无法连接 SAP：确认已登录 SAP，并且 SAP GUI Scripting 已启用。
- 控件找不到或超时：确认 SAP 版本、界面布局、事务码界面字段与脚本中控件路径一致。
- Excel 文件不存在：修改脚本底部的 `EXCEL_PATH` 或 `excel_file` 为实际文件路径。
- 中文显示异常：请使用 UTF-8 保存脚本，避免用不支持 UTF-8 的编辑器覆盖文件。
- 自动化执行前请先备份 Excel 原始数据，脚本会直接回写处理结果。
