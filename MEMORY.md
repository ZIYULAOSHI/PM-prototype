# MEMORY：量化王智能客服项目 - 经验总结

## 项目背景

在 Windows 系统上，通过 Kiro IDE 对 `智能客服/智能问答.xlsx` 进行读取分析和写入操作。该 Excel 文件包含近 5000 条用户群聊记录（Sheet1）、意图分类体系（Sheet2），需要在 Sheet3 中生成标准化智能问答数据。

---

## 一、环境问题

### 1. Python / Node.js 不可用

**问题：** 系统上 `python3`、`python`、`node` 命令均不可用。`where.exe python` 返回的是 Windows Store 占位符路径（`WindowsApps\python.exe`），实际并未安装 Python。

**解决：** 放弃 Python/Node 方案，改用 **PowerShell COM 对象**（`New-Object -ComObject Excel.Application`）直接操作 Excel，这是 Windows 系统上无需额外安装依赖的可靠方式。

---

## 二、读取 Excel 的问题与解决

### 2. 中文路径导致 PowerShell 脚本文件执行失败

**问题：** 将包含中文字符串的代码写入 `.ps1` 文件后执行，中文会变成乱码（如 `鏅€氳亰澶?`），导致字符串匹配失败、语法解析错误（`字符串缺少终止符`）。

**根因：** Kiro 的 `fsWrite` 工具写入文件时使用的编码（UTF-8 无 BOM）与 PowerShell 读取 `.ps1` 文件时期望的编码不一致。PowerShell 5.x 默认使用系统区域设置编码（GBK/GB2312）来读取脚本文件。

**解决：** 
- **方案 A（推荐）：** 不使用 `.ps1` 脚本文件，直接在 PowerShell 命令行中执行中文代码。命令行输入不经过文件编码转换，中文可以正确传递。
- **方案 B：** 如果必须用脚本文件，避免在脚本中硬编码中文字符串，改用 Unicode 转义或从 Excel 中动态读取。

### 3. 逐行遍历大量数据导致超时

**问题：** 用 `for` 循环逐行读取 4972 行数据（`$ws.Cells.Item($r, 1).Text`），COM 对象的逐单元格访问非常慢，导致命令超时。

**解决：** 使用 **Range 批量读取**，一次性将整列数据读入内存数组：

```powershell
$rangeA = $ws.Range("A2:A$maxRow")
$rangeB = $ws.Range("B2:B$maxRow")
$valA = $rangeA.Value2
$valB = $rangeB.Value2
# 然后在内存中遍历 $valA[$i, 1] 和 $valB[$i, 1]
```

这种方式将读取时间从几分钟缩短到几秒。

### 4. 命令行过长导致执行失败

**问题：** 将所有逻辑写在一行命令中，命令过长时 PowerShell 终端会截断或无法正确解析。

**解决：** 使用多行命令格式（PowerShell 支持在终端中直接输入多行代码），或将逻辑拆分为多个独立的命令块分批执行。

---

## 三、写入 Excel 的问题与解决

### 5. `.Value2` 赋值方式在某些场景下不生效

**问题：** 使用 `$ws.Cells.Item(1,1).Value2 = "xxx"` 写入数据，在当前会话中可以读取到，但保存后重新打开文件数据丢失。

**解决：** 改用直接赋值语法：

```powershell
# 不可靠（某些场景下不持久化）：
$ws.Cells.Item(1,1).Value2 = "xxx"

# 可靠（推荐）：
$ws.Cells.Item(1,1) = "xxx"
```

### 6. `$wb.Save()` 保存后数据丢失（核心问题）

**问题：** 这是整个过程中最棘手的问题。写入数据后调用 `$wb.Save()`，在当前 COM 会话中验证数据存在，但关闭后重新打开文件数据全部丢失。

**根因：** 原始文件 `智能问答.xlsx` 存在 `~$智能问答.xlsx` 锁定文件，说明该文件可能被其他 Excel 进程或用户打开。在文件被锁定的情况下，`$wb.Save()` 可能静默失败（不报错但实际未写入磁盘）。

**解决：** 使用 `SaveAs` 保存到新文件，绕过文件锁定：

```powershell
# 不可靠（文件被锁定时静默失败）：
$wb.Save()

# 可靠（保存到新文件）：
$newPath = "C:\...\智能问答_output.xlsx"
$wb.SaveAs($newPath, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
```

**预防措施：**
- 写入前先清理残留 Excel 进程：`Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force`
- 删除锁定文件：`Remove-Item "~$*.xlsx" -Force`
- 写入后立即重新打开文件验证数据是否持久化

### 7. 批量写入的高效方式

**问题：** 逐单元格写入（`$ws.Cells.Item($r,$c) = "xxx"`）在数据量大时非常慢。

**解决：** 使用 **二维数组 + Range 批量写入**：

```powershell
$d = New-Object 'object[,]' 10,7
$d[0,0] = "值1"; $d[0,1] = "值2"; ...
$d[1,0] = "值3"; $d[1,1] = "值4"; ...

$ws.Range("A1:G10").Value2 = $d
```

这种方式一次性写入整个区域，速度比逐单元格快几十倍。

### 8. 分批写入避免命令超时

**问题：** 一次性写入 40+ 行 × 7 列的数据，命令文本量巨大，容易超时。

**解决：** 将数据分为 4 批（每批约 10 行），每批独立执行：
1. 打开文件 → 写入 10 行 → 保存 → 关闭
2. 打开文件 → 写入下 10 行 → 保存 → 关闭
3. 重复直到全部写入

每批之间加 `Start-Sleep -Seconds 1` 确保 COM 对象完全释放。

---

## 四、最佳实践总结

### 读取 Excel

```powershell
# 1. 创建 COM 对象
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 2. 打开文件（用动态路径避免中文硬编码问题）
$dirs = Get-ChildItem -Directory
$f = ($dirs | ForEach-Object { 
    Get-ChildItem -Path $_.FullName -Filter "*.xlsx" | 
    Where-Object { $_.Name -notlike '~$*' } 
}) | Select-Object -First 1
$wb = $excel.Workbooks.Open($f.FullName)

# 3. 批量读取数据
$ws = $wb.Sheets.Item(1)
$maxRow = $ws.UsedRange.Rows.Count
$values = $ws.Range("A2:B$maxRow").Value2

# 4. 在内存中处理
for ($i = 1; $i -le $values.GetLength(0); $i++) {
    $colA = $values[$i, 1]
    $colB = $values[$i, 2]
}

# 5. 清理
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
```

### 写入 Excel

```powershell
# 1. 先清理残留进程
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

# 2. 打开文件
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($filePath)
$ws = $wb.Sheets.Item(3)

# 3. 用二维数组批量写入
$d = New-Object 'object[,]' 行数,列数
$d[0,0] = "数据"; ...
$ws.Range("A1:G10").Value2 = $d

# 4. SaveAs 到新文件（避免锁定问题）
$wb.SaveAs($newPath, 51)

# 5. 清理
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
```

### 关键注意事项

| 问题 | 解决方案 |
|------|---------|
| 中文在 .ps1 文件中乱码 | 直接在命令行执行，不写入脚本文件 |
| 逐单元格读写太慢 | 用 Range + 二维数组批量操作 |
| Save() 静默失败 | 用 SaveAs() 保存到新文件 |
| 文件被锁定 | 先 Stop-Process excel，删除 ~$ 锁定文件 |
| 命令太长超时 | 分批执行，每批 10 行左右 |
| COM 对象未释放 | 必须调用 ReleaseComObject，批次间加 Sleep |

---

## 五、项目文件说明

| 文件 | 说明 |
|------|------|
| `智能客服/智能问答.xlsx` | 原始数据文件（Sheet1: 用户对话, Sheet2: 意图分类体系） |
| `智能客服/智能问答_output.xlsx` | 输出文件（在原文件基础上，Sheet3 填入了 40 条标准化问答数据） |

Sheet3 结构：用户意图分类 / 适用条件 / 标准提问 / 相似问法 / 解决方案 / 关联问题 / 错误回答
