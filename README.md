# ReadBom

SolidWorks BOM 读取与属性编辑工具，包含 WPF 桌面主程序和 SolidWorks COM Add-in 两部分，通过本地 HTTP 通信。

## 项目结构

```
readbom/
├── readbom/                    # WPF 主程序 (.NET 9.0)
│   ├── MainWindow.xaml/.cs     # 主界面
│   ├── readbom.csproj
│   └── property-mapping.txt   # TXT 属性映射配置
├── ReadBom.SwAddin/            # SolidWorks Add-in (.NET Framework 4.8)
│   ├── SwAddin.cs             # 插件入口，启动 HTTP 服务
│   ├── AddinHttpServer.cs     # HTTP 命令处理
│   └── CommandRequest.cs      # 请求 DTO
└── docs/
    └── sw-api-reference.md    # SolidWorks API 参考
```

## 架构

```
┌──────────────┐    HTTP (127.0.0.1:32127)    ┌──────────────────┐
│  readbom     │ ◄──────────────────────────► │ ReadBom.SwAddin  │
│  (WPF 桌面)   │    JSON 命令/响应             │  (SW 进程内插件)  │
└──────────────┘                               └──────────────────┘
```

- **主程序**负责 UI、CSV 解析、数据校验、Excel 导出、文件复制等业务逻辑
- **Add-in**只负责必须在 SolidWorks 进程内完成的操作（读取 BOM、读写属性、打开文档等）

## 功能

- 按 BOM 读取装配体零件清单，显示类型、文件名、配置、数量、材质、工程图状态
- 支持自定义属性列的动态读取与显示
- 表格内编辑属性值，批量保存回 SolidWorks
- 复制文件到主装配体目录并重装引用
- 计算钣金件下料尺寸
- 表头筛选、复制列、填充列、查找替换
- 缩略图预览
- 导出 CSV / Excel
- 右键菜单快速打开零件、装配体、工程图

## 构建

**主程序：**

```powershell
dotnet build readbom\readbom.csproj
```

**Add-in：**

```powershell
dotnet build ReadBom.SwAddin\ReadBom.SwAddin.csproj
```

## 注册 Add-in

使用 64 位 RegAsm 注册 COM Add-in：

```powershell
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe ReadBom.SwAddin.dll /codebase
```

取消注册：

```powershell
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe ReadBom.SwAddin.dll /unregister
```

注册后启动 SolidWorks，插件会自动加载并监听 `http://127.0.0.1:32127/`。

## 使用方式

1. 启动 SolidWorks，确认 Add-in 已加载
2. 打开主程序 `readbom.exe`
3. 点击"连接 SW"检查连接状态
4. 打开目标装配体，选择读取方式，点击"读取属性/BOM"
5. 在表格中编辑属性，点击"保存到 SW"写入
6. 使用工具栏或右键菜单进行导出、筛选等操作

## 依赖

- SolidWorks 2022+ (64-bit)
- .NET 9.0 Desktop Runtime（主程序）
- .NET Framework 4.8（Add-in）

## Add-in HTTP 接口

详见 [ReadBom.SwAddin/README.md](ReadBom.SwAddin/README.md)

## SW API 参考

详见 [docs/sw-api-reference.md](docs/sw-api-reference.md)