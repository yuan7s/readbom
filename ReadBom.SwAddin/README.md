# ReadBom.SwAddin

SolidWorks COM Add-in，用作 `readbom` 主程序和 SolidWorks 之间的本地 HTTP 数据通道。

Add-in 只负责必须在 SolidWorks 进程内完成的读取和导出动作。CSV 解析、文本转换、材料判断、工程图判断、完整路径回填、校验、计算和 UI 展示都由主程序完成。

## 职责边界

Add-in 负责：

- 在 SolidWorks 中加载并监听本地 HTTP：`http://127.0.0.1:32127/`
- 获取当前活动文档基础信息。
- 读取当前文档属性。
- 写入单个属性。
- 执行重建、保存等简单 SolidWorks 命令。
- 为 BOM 准备或复用 `.sldbomtbt` 模板。
- 插入隐藏 BOM 表。
- 调用 `SaveAsText2` 导出 BOM CSV。
- 执行 `open-document`：打开指定零件、装配体或工程图，并激活 SolidWorks 窗口。
- 执行 `related-files`：通过 SolidWorks 获取主装配体相关文件列表。
- 读取 CSV 原始字节并以 Base64 返回主程序。
- 删除临时 BOM 表和临时 CSV 文件。

Add-in 不负责：

- CSV 编码识别和文本解析。
- BOM 行对象构建。
- 材料值公式清理，例如 `$PRP:材质`。
- 材料是否未设置判断。
- 主程序 UI 操作、筛选、校验和表格编辑。
- 完整路径回填。
- 工程图是否存在判断。
- 复制文件、重装引用、下料尺寸计算。
- Excel 导出和表格 UI 处理。

## HTTP 接口

健康检查：

```http
GET http://127.0.0.1:32127/health
```

执行命令：

```http
POST http://127.0.0.1:32127/command
Content-Type: application/json

{"Command":"active-document"}
```

响应外层格式：

```json
{
  "ok": true,
  "data": {}
}
```

失败时：

```json
{
  "ok": false,
  "error": "错误信息"
}
```

## 支持命令

- `ping`
- `active-document`
- `list-properties`
- `read-bom`
- `set-property`
- `rebuild`
- `save`

## read-bom

请求示例：

```json
{
  "Command": "read-bom",
  "PropertyNames": ["物料编码", "零件图号", "材料"],
  "GroupByConfig": true,
  "SkipVirtual": true
}
```

当前装配体路径会走 BOM CSV 通道。Add-in 返回：

```json
{
  "tableCsv": {
    "csvBase64": "base64 encoded csv bytes",
    "csvByteCount": 12345,
    "separator": ",",
    "propertyNames": ["物料编码", "零件图号", "材料"],
    "mainPath": "F:\\Project\\Main.SLDASM",
    "mainConfiguration": "默认",
    "rowCount": 135
  }
}
```

说明：

- `csvBase64` 是 `SaveAsText2` 导出的原始 CSV 字节，不在 Add-in 中做编码转换。
- `separator` 当前为英文逗号 `,`。
- `propertyNames` 是本次主程序请求的 TXT 属性列。
- `mainPath` 和 `mainConfiguration` 只作为主程序构建主装配体行和后续路径解析的上下文。
- CSV 解析、字段映射和数据清洗全部在主程序中完成。

## BOM 模板

`InsertBomTable3` 使用 SolidWorks BOM 模板 `.sldbomtbt`，不支持用 `.xls/.xlsx/.xsl` 作为插入模板。

Add-in 会根据请求的属性列准备专用 BOM 模板：

1. 检查 DLL 目录下的模板缓存。
2. 缓存不存在时创建临时空装配体。
3. 在临时空装配体中插入基础 BOM。
4. 添加 `SW材料` 和 TXT 属性列。
5. `SaveAsTemplate` 保存为专用 `.sldbomtbt`。
6. 关闭临时空装配体。
7. 当前主装配体只使用专用模板插入 BOM 并导出 CSV。

模板缓存位置：

```text
ReadBom.SwAddin\bin\Debug\net48\BomTemplates
```

临时 CSV 导出位置：

```text
ReadBom.SwAddin\bin\Debug\net48\BomExports
```

CSV 读取后会尝试删除临时文件。

## 日志

日志写到 Add-in DLL 所在目录：

```text
ReadBom.SwAddin\bin\Debug\net48\SwAddin.log
```

典型 BOM 日志：

```text
Command start: read-bom
ReadBom timing: prepared BOM template cache hit ...
ReadBom timing: InsertBomTable3 only elapsed=...
ReadBom timing: SaveAsText2 CSV saved=True ...
ReadBom timing: read CSV file bytes=...
ReadBom timing: delete hidden BOM elapsed=...
Command response written: read-bom ...
```

## 构建

项目目标框架：

```text
.NET Framework 4.8
PlatformTarget: x64
```

构建：

```powershell
dotnet build ReadBom.SwAddin\ReadBom.SwAddin.csproj
```

## 注册

需要使用 64 位 .NET Framework RegAsm 注册 COM Add-in。

注册：

```powershell
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe ReadBom.SwAddin.dll /codebase
```

取消注册：

```powershell
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe ReadBom.SwAddin.dll /unregister
```

如果 SolidWorks 中看不到插件，检查：

- 是否注册的是当前构建输出目录中的 DLL。
- SolidWorks 和注册命令是否使用相同权限级别。
- `HKCU\SOFTWARE\SolidWorks\AddIns` 和 `HKLM\SOFTWARE\SolidWorks\AddIns` 中是否存在插件 GUID。
- `SwAddin.log` 是否有 `ConnectToSW called`。

## 开发注意

- 修改 Add-in 后需要重新加载插件或重启 SolidWorks。
- 如果 DLL 被 SolidWorks 占用，先关闭 SolidWorks 再构建正式输出目录。
- Add-in 不应引入 UI 或业务计算逻辑；新增业务规则优先放到主程序。
