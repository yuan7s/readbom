# SolidWorks API 参考

整理自 `chajian2` 项目（`C:\Users\yuan7\RiderProjects\chajian2`）的 SW API 使用方式，供 `readbom` 项目参考。

## COM 互操作引用

需要引用两个互操作程序集：

| 程序集 | 说明 |
|--------|------|
| `Interop.SldWorks.dll` | 主要 SolidWorks API 类型（SldWorks 命名空间） |
| `Interop.SwConst.dll` | SolidWorks 枚举常量（SwConst 命名空间） |

`readbom` 项目使用 NuGet 包引用，`chajian2` 使用本地 `lib\` 目录下的 DLL 并启用 `EmbedInteropTypes=True`。

---

## 一、SldWorks 应用对象

### 获取 SW 实例

**方式 1：COM GetActiveObject（chajian2 外部应用方式）**

```vb
' 获取当前运行的 SW 实例
Dim swApp As SldWorks.SldWorks = Marshal.GetActiveObject("SldWorks.Application")
```

**方式 2：Running Object Table 枚举（多实例时）**

通过 ROT 按 moniker `solidworks_pid_<PID>` 查找特定进程：

```vb
' 使用 ole32.dll 的 GetRunningObjectTable / CreateBindCtx
' 枚举 ROT 条目，匹配 moniker 名称前缀 "solidworks_pid_"
```

**方式 3：Add-in ConnectToSW（readbom 插件方式）**

```csharp
// 插件加载时 SW 自动传入
public bool ConnectToSW(object thisSw, int cookie)
{
    _swApp = (SldWorks.SldWorks)thisSw;
}
```

### SldWorks 对象常用属性/方法

| 成员 | 说明 | chajian2 | readbom |
|------|------|----------|---------|
| `.ActiveDoc` | 当前活动文档（返回 ModelDoc2） | 全部文件 | AddinHttpServer |
| `.GetProcessID()` | 获取 SW 进程 ID | Form1.vb | — |
| `.Frame()` | 获取 SW 主窗口框架对象 | Form1.vb | AddinHttpServer |
| `.Frame().GetHWnd()` | 获取 SW 主窗口句柄 | Form1.vb | AddinHttpServer |
| `.Visible` | 设置 SW 窗口可见性 | — | AddinHttpServer |
| `.OpenDoc6(path, type, options, config, errors, warnings)` | 打开文档 | Form1/Rename/Coding | AddinHttpServer |
| `.GetOpenDocumentByName(path)` | 按路径查找已打开文档 | RenameWindow | AddinHttpServer |
| `.ActivateDoc3(title, ...)` | 激活已打开文档 | RenameWindow | AddinHttpServer |
| `.CloseDoc(title)` | 关闭文档 | — | AddinHttpServer |
| `.GetFirstDocument()` | 获取第一个文档（遍历用） | — | AddinHttpServer |
| `.GetDocumentDependencies2(path, ...)` | 获取文档依赖文件列表 | — | AddinHttpServer |
| `.NewDocument(template, ...)` | 从模板新建文档 | — | AddinHttpServer |
| `.NewAssembly()` | 新建空装配体 | — | AddinHttpServer |
| `.GetUserPreferenceStringValue(...)` | 获取用户偏好字符串 | — | AddinHttpServer |
| `.SetAddinCallbackInfo2(...)` | 设置插件回调 | — | SwAddin |
| `.SendMsgToUser2(msg, icon, buttons)` | 在 SW 中显示消息 | Form1.vb | — |

### SW 事件（chajian2 使用 WithEvents 订阅）

| 事件 | 触发时机 | 使用文件 |
|------|----------|----------|
| `ActiveDocChangeNotify` | 活动文档切换 | Form1, RenameWindow, PropertyOverlayWindow |
| `ActiveModelDocChangeNotify` | 活动模型文档变更 | Form1, RenameWindow, PropertyOverlayWindow |
| `FileCloseNotify` | 文件关闭 | Form1 |
| `PartDoc.NewSelectionNotify` | 零件中选择变更 | Form1, RenameWindow, PropertyOverlayWindow |
| `AssemblyDoc.NewSelectionNotify` | 装配体中选择变更 | Form1, RenameWindow, PropertyOverlayWindow |
| `DrawingDoc.NewSelectionNotify` | 工程图中选择变更 | Form1, RenameWindow, PropertyOverlayWindow |

```vb
' chajian2 事件订阅模式
Private WithEvents _swApp As SldWorks.SldWorks

Private Sub _swApp_ActiveDocChangeNotify() Handles _swApp.ActiveDocChangeNotify
    ' 处理文档切换
End Sub
```

---

## 二、文档操作（ModelDoc2）

`ModelDoc2` 是所有文档类型的基类。

### 基本信息

| 成员 | 说明 | 返回值 |
|------|------|--------|
| `.GetType()` | 文档类型 | `swDocPART` / `swDocASSEMBLY` / `swDocDRAWING` |
| `.GetPathName()` | 文档保存路径 | string |
| `.GetTitle()` | 文档显示标题 | string |
| `.GetNext()` | 获取下一个文档（遍历） | ModelDoc2 |

### 配置管理

| 成员 | 说明 |
|------|------|
| `.ConfigurationManager.ActiveConfiguration.Name` | 当前活动配置名称 |
| `.GetActiveConfiguration()` | 获取活动配置对象 |
| `.GetConfigurationByName(name)` | 按名称获取配置 |
| `.GetConfigurationNames()` | 获取所有配置名称数组 |
| `.ShowConfiguration2(name)` | 切换配置 |

### 选择管理

| 成员 | 说明 |
|------|------|
| `.SelectionManager` | 获取选择管理器 |
| `.ClearSelection2(force)` | 清除选择集 |

### 保存与重建

| 成员 | 说明 |
|------|------|
| `.Save3(options, errors, warnings)` | 保存文档 |
| `.SaveAs3(...)` | 另存为 |
| `.Extension.SaveAs(...)` | 通过 Extension 另存 |
| `.EditRebuild3()` | 重建模型 |
| `.ForceRebuild3(...)` | 强制重建 |
| `.SetSaveFlag()` | 标记文档为脏 |

### 自定义属性

| 成员 | 说明 |
|------|------|
| `.Extension.get_CustomPropertyManager(configName)` | 获取属性管理器（传空字符串获取文档级属性） |
| `.GetCustomInfoNames()` | 获取自定义属性名称列表（文档级） |
| `.GetCustomInfoNames2(config)` | 获取配置级属性名称列表 |
| `.GetCustomInfoValue(config, name)` | 获取属性值 |
| `.DeleteCustomInfo(name)` | 删除文档级属性 |
| `.DeleteCustomInfo2(config, name)` | 删除配置级属性 |

### 用户偏好

| 成员 | 说明 |
|------|------|
| `.Extension.SetUserPreferenceInteger(key, value)` | 设置用户偏好整数 |
| `.GetUserPreferenceStringValue(key)` | 获取用户偏好字符串 |

### BOM 表操作（readbom 特有）

| 成员 | 说明 |
|------|------|
| `.Extension.InsertBomTable3(template, x, y, type, config, ...)` | 插入 BOM 表 |
| `.Extension.DeleteSelection2(options)` | 删除选择（用于删除隐藏 BOM 表） |
| `.Extension.SelectByID2(name, type, x, y, z, ...)` | 按 ID 选择对象 |

---

## 三、零件文档（PartDoc）

```csharp
var partDoc = model as PartDoc;
```

| 成员 | 说明 | chajian2 | readbom |
|------|------|----------|---------|
| `.GetPartBox(includeRefPlanes)` | 获取零件包围盒 | RenameWindow | — |
| `.GetMaterialPropertyName2(config, out db)` | 获取配置材质名称 | — | AddinHttpServer |
| `.GetMaterialPropertyName(out db)` | 获取材质名称（旧版） | — | AddinHttpServer |
| `.MaterialUserName` | 材质用户名称属性 | — | AddinHttpServer |
| `.MaterialIdName` | 材质 ID 名称属性 | — | AddinHttpServer |

---

## 四、装配体文档（AssemblyDoc）

```csharp
var assembly = model as AssemblyDoc;
```

| 成员 | 说明 | chajian2 | readbom |
|------|------|----------|---------|
| `.GetComponents(topLevel)` | 获取装配体组件列表 | Form1/Coding | AddinHttpServer |
| `.GetBox(options)` | 获取装配体包围盒 | RenameWindow | AddinHttpServer |
| `.ResolveAllLightWeightComponents()` | 还原所有轻化组件 | Coding | — |
| `.ReorderComponents(names, where, target)` | 在设计树中排序组件 | Form1 | — |

---

## 五、工程图文档（DrawingDoc）

```csharp
var drawing = model as DrawingDoc;
```

| 成员 | 说明 | chajian2 |
|------|------|----------|
| `.ReplaceReferencedDocument(...)` | 替换工程图引用的文档路径 | RenameWindow |

### 视图操作

| 成员 | 说明 | chajian2 |
|------|------|----------|
| `View.Name` | 视图名称 | Form1.vb |
| `View.Angle` | 视图旋转角度 | Form1.vb |

---

## 六、组件操作（Component2）

```csharp
var component = item as Component2;
```

| 成员 | 说明 | chajian2 | readbom |
|------|------|----------|---------|
| `.GetModelDoc2()` | 获取组件引用的文档 | 全部文件 | AddinHttpServer |
| `.GetModelDoc()` | 获取组件文档（旧版） | Form1/Coding | — |
| `.GetPathName()` | 获取组件文件路径 | Form1/Rename/Coding | AddinHttpServer |
| `.Name2` | 组件实例名称 | Form1 | — |
| `.Name` | 组件名称 | Coding | — |
| `.ReferencedConfiguration` | 组件引用的配置名 | Form1/Coding | AddinHttpServer |
| `.GetChildren()` | 获取子组件 | Form1/Coding | — |
| `.IsSuppressed()` | 是否被压缩 | Form1 | AddinHttpServer |
| `.GetSuppression()` | 获取压缩状态 | — | AddinHttpServer |
| `.IsEnvelope()` | 是否为封套组件 | Form1 | — |
| `.IsVirtual` | 是否为虚拟组件 | — | AddinHttpServer |
| `.ReplaceReference(...)` | 替换组件引用路径 | RenameWindow | — |

---

## 七、选择管理器（SelectionMgr）

```csharp
var selMgr = model.SelectionManager;
```

| 成员 | 说明 |
|------|------|
| `.GetSelectedObjectCount2(mark)` | 获取选中对象数量（-1 = 所有标记） |
| `.GetSelectedObject6(index, mark)` | 获取指定索引的选中对象 |
| `.GetSelectedObjectsComponent(index)` | 获取选中对象的组件引用 |

---

## 八、特征管理器与设计树（chajian2 特有）

### FeatureManager

```csharp
var featMgr = model.FeatureManager;
```

| 成员 | 说明 |
|------|------|
| `.GetFeatures(topLevel)` | 获取所有特征 |
| `.HideComponentSingleConfigurationOrDisplayStateNames` | 隐藏配置/显示状态名称 |
| `.ShowComponentConfigurationNames` | 显示配置名称 |
| `.ShowComponentConfigurationDescriptions` | 显示配置描述 |
| `.ShowDisplayStateNames` | 显示显示状态名称 |
| `.SetComponentIdentifiers(...)` | 设置组件标识符显示模式 |

### Feature（特征对象）

| 成员 | 说明 |
|------|------|
| `.GetSpecificFeature2()` | 获取特征的底层对象（如 Component2） |
| `.GetTypeName2()` | 获取特征类型字符串 |
| `.GetFirstSubFeature()` | 获取文件夹内第一个子特征 |
| `.GetNextSubFeature()` | 获取文件夹内下一个子特征 |
| `.Select2(...)` | 选中特征 |

---

## 九、自定义属性管理器（CustomPropertyManager）

```csharp
var manager = model.Extension.get_CustomPropertyManager(configName);
```

| 成员 | 说明 | chajian2 | readbom |
|------|------|----------|---------|
| `.Add3(name, type, value, option)` | 添加/替换属性（推荐） | Form1/Coding/Rename | AddinHttpServer |
| `.Add2(name, type, value)` | 添加属性（旧版） | — | AddinHttpServer |
| `.Set2(name, value)` | 设置属性值 | — | AddinHttpServer |
| `.Set(name, value)` | 设置属性值（更旧版） | — | AddinHttpServer |
| `.GetAll3(names, types, values, status, link)` | 获取所有属性（含链接信息） | — | AddinHttpServer |
| `.GetAll2(names, types, values, status)` | 获取所有属性（旧版） | — | AddinHttpServer |

**write 回退链（readbom）：** `Add3` → `Set2` → `Set` → `Add2`

---

## 十、BOM 表操作（readbom 特有）

### BomTableAnnotation

```csharp
var bomTable = model.Extension.InsertBomTable3(...) as BomTableAnnotation;
```

| 成员 | 说明 |
|------|------|
| `.BomFeature.GetFeature()` | 获取 BOM 特征对象（用于删除） |
| `.SetColumnCustomProperty(colIndex, propName)` | 设置列为自定义属性列 |

### TableAnnotation

```csharp
var table = (TableAnnotation)bomTable;
```

| 成员 | 说明 |
|------|------|
| `.RowCount` | 行数 |
| `.ColumnCount` | 列数 |
| `.InsertColumn2(position, index, title, widthStyle)` | 插入列 |
| `.SetColumnType2(colIndex, type, ...)` | 设置列类型 |
| `.SetColumnType3(colIndex, type, ..., propName)` | 设置列类型（带属性名） |
| `.SetColumnTitle2(colIndex, title, ...)` | 设置列标题 |
| `.SaveAsText2(path, separator, ...)` | 导出为文本/CSV |
| `.SaveAsTemplate(path)` | 保存为 BOM 模板 (.sldbomtbt) |

### BOM 模板路径

```
C:\Program Files\SOLIDWORKS\lang\chinese-simplified\bom-standard.sldbomtbt
C:\Program Files\SOLIDWORKS\lang\english\bom-standard.sldbomtbt
```

---

## 十一、SwConst 枚举常量

### swDocumentTypes_e

| 常量 | 值 | 说明 |
|------|-----|------|
| `swDocPART` | — | 零件文档 |
| `swDocASSEMBLY` | — | 装配体文档 |
| `swDocDRAWING` | — | 工程图文档 |

### swOpenDocOptions_e

| 常量 | 说明 |
|------|------|
| `swOpenDocOptions_Silent` | 静默打开（不显示 UI） |

### swSaveAsVersion_e

| 常量 | 说明 |
|------|------|
| `swSaveAsCurrentVersion` | 保存为当前版本 |

### swSaveAsOptions_e

| 常量 | 说明 |
|------|------|
| `swSaveAsOptions_Silent` | 静默保存 |
| `swSaveAsOptions_Copy` | 保存副本 |

### swCustomInfoType_e

| 常量 | 说明 |
|------|------|
| `swCustomInfoText` | 文本类型属性 |

### swCustomPropertyAddOption_e

| 常量 | 值 | 说明 |
|------|-----|------|
| `swCustomPropertyDeleteAndAdd` | 2 | 删除后添加 |
| `swCustomPropertyReplaceValue` | — | 替换值 |

### swBomType_e

| 常量 | 说明 |
|------|------|
| `swBomType_Indented` | 缩进式 BOM |

### swNumberingType_e

| 常量 | 说明 |
|------|------|
| `swNumberingType_Detailed` | 详细编号 |

### swTableItemInsertPosition_e

| 常量 | 说明 |
|------|------|
| `swTableItemInsertPosition_Last` | 插入到末尾 |

### swTableColumnTypes_e

| 常量 | 说明 |
|------|------|
| `swBomTableColumnType_CustomProperty` | 自定义属性列 |

### swInsertTableColumnWidthStyle_e

| 常量 | 说明 |
|------|------|
| `swInsertColumn_DefaultWidth` | 默认列宽 |

### swDeleteSelectionOptions_e

| 常量 | 说明 |
|------|------|
| `swDelete_Absorbed` | 删除被吸收的特征 |

### swBoundingBoxOptions_e

| 常量 | 说明 |
|------|------|
| `swBoundingBoxIncludeRefPlanes` | 包围盒包含参考平面 |

### swRebuildOnActivation_e

| 常量 | 说明 |
|------|------|
| `swUserDecision` | 用户决定 |

### swReorderComponentsWhere_e

| 常量 | 说明 |
|------|------|
| `swReorderComponents_After` | 排序到目标之后 |

### swMessageBoxIcon_e

| 常量 | 说明 |
|------|------|
| `swMbInformation` | 信息图标 |

### swMessageBoxBtn_e

| 常量 | 说明 |
|------|------|
| `swMbOk` | 确定按钮 |

### swUserPreferenceIntegerValue_e

| 常量 | 说明 |
|------|------|
| `swDetailingDimensionStandard` | 尺寸标准偏好 |

### swDetailingStandard_e

| 常量 | 说明 |
|------|------|
| `swDetailingStandardISO` | ISO 标准 |

### swUserPreferenceStringValue_e

| 常量 | 说明 |
|------|------|
| `swDefaultTemplateAssembly` | 默认装配体模板路径 |

---

## 十二、文档类型判断（通过扩展名）

```csharp
static int GetDocumentTypeFromPath(string path)
{
    var ext = Path.GetExtension(path);
    if (ext.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)) return (int)swDocumentTypes_e.swDocPART;
    if (ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return (int)swDocumentTypes_e.swDocASSEMBLY;
    if (ext.Equals(".slddrw", StringComparison.OrdinalIgnoreCase)) return (int)swDocumentTypes_e.swDocDRAWING;
    return 0;
}
```

---

## 十三、chajian2 关键架构模式

### 外部应用 vs 插件

| | chajian2 | readbom |
|---|---|---|
| 运行方式 | 独立 .exe 外部应用 | COM Add-in 插件 |
| SW 连接 | `Marshal.GetActiveObject()` + ROT 枚举 | `ConnectToSW` 自动注入 |
| 进程生命周期 | 用户启动/关闭 exe | SW 启动/关闭时加载/卸载 |
| 多实例支持 | ROT 按 PID 匹配 | 每个 SW 进程加载一个插件副本 |
| COM 清理 | `Marshal.ReleaseComObject` | 无需手动管理 |

### ROT 多实例发现（chajian2 特有）

```vb
' 通过 Running Object Table 枚举所有 SW 实例
' moniker 格式: solidworks_pid_<进程ID>
' 使用 ole32.dll: GetRunningObjectTable, CreateBindCtx
```

### 文档事件附加模式（chajian2）

```vb
' 通过 AddHandler 动态订阅文档级事件
Private Sub AttachDocEvents(model As ModelDoc2)
    If TypeOf model Is PartDoc Then
        AddHandler CType(model, PartDoc).NewSelectionNotify, AddressOf OnSelectionChange
    ElseIf TypeOf model Is AssemblyDoc Then
        AddHandler CType(model, AssemblyDoc).NewSelectionNotify, AddressOf OnSelectionChange
    ElseIf TypeOf model Is DrawingDoc Then
        AddHandler CType(model, DrawingDoc).NewSelectionNotify, AddressOf OnSelectionChange
    End If
End Sub
```

---

## 十四、readbom 特有模式

### HTTP 命令分发

```
POST /command  →  ExecuteCommand(request)  →  switch(command):
  ping               →  { message: "pong" }
  active-document    →  GetActiveDocumentInfo()
  list-properties    →  ListProperties(request)
  read-bom           →  ReadBom(request)
  open-document      →  OpenDocument(request)
  related-files      →  GetRelatedFiles(request)
  save-properties-batch → SavePropertiesBatch(request)
  calculate-blank-size  → CalculateBlankSizeBatch(request)
  set-property       →  SetProperty(request)
  rebuild            →  Rebuild()
  save               →  Save()
```

### 属性读取三级来源

```
配置属性 (get_CustomPropertyManager(configName))
  → 文档属性 (get_CustomPropertyManager(""))
    → 空字符串
```

### 材质读取回退链

```
GetMaterialPropertyName2(config)
  → GetMaterialPropertyName2(activeConfig)
    → GetMaterialPropertyName2("")
      → GetMaterialPropertyName2("Default")
        → GetMaterialPropertyName (无参)
          → MaterialUserName
            → MaterialIdName
              → ""
```

### BOM 模板缓存策略

1. 根据请求的属性列计算签名（SHA256 前 16 位）
2. 检查 `BomTemplates\readbom-{hash}.sldbomtbt` 是否存在
3. 缓存未命中时：创建临时空装配体 → 插入基础 BOM → 添加 SW材料 + 属性列 → SaveAsTemplate → 关闭临时装配体
4. 后续请求命中缓存直接使用

### 隐藏 BOM 表生命周期

```
InsertBomTable3 → 导出 CSV → 读取 CSV 字节 → 删除 BOM 表特征 → 删除临时 CSV 文件
```

---

## 十五、常用代码片段

### 安全调用（忽略异常）

```csharp
static T Safe<T>(Func<T> func)
{
    try { return func(); }
    catch { return default; }
}
```

### 获取文档类型标签

```csharp
static string GetDocumentTypeLabel(string path)
{
    var ext = Path.GetExtension(path ?? "");
    if (ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return "装配体";
    if (ext.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)) return "零件";
    if (ext.Equals(".slddrw", StringComparison.OrdinalIgnoreCase)) return "工程图";
    return "未知";
}
```

### 写入属性（带回退）

```csharp
static void WriteProperty(CustomPropertyManager manager, string name, string value)
{
    // 依次尝试 Add3 → Set2 → Set → Add2
    manager.Add3(name, (int)swCustomInfoType_e.swCustomInfoText, value,
        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
}
```

### 读取所有属性

```csharp
Dictionary<string, string> ReadAllProperties(CustomPropertyManager manager)
{
    object namesObj = null, typesObj = null, valuesObj = null, statusObj = null, linkObj = null;
    try { manager.GetAll3(ref namesObj, ref typesObj, ref valuesObj, ref statusObj, ref linkObj); }
    catch { manager.GetAll2(ref namesObj, ref typesObj, ref valuesObj, ref statusObj); }
    // namesObj, valuesObj 转为 string[] 或 object[] 遍历
}
```

### 遍历已打开文档

```csharp
var model = _swApp.GetFirstDocument() as ModelDoc2;
while (model != null)
{
    var path = Safe(() => model.GetPathName());
    // ...
    model = model.GetNext() as ModelDoc2;
}
```