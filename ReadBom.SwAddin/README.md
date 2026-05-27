# ReadBom.SwAddin

SolidWorks Add-in HTTP 通讯桥。

插件加载到 SolidWorks 后会监听：

`http://127.0.0.1:32127/`

主程序可以通过 HTTP 调用插件，让插件在 SolidWorks 进程内执行简单命令并返回 JSON。

## 接口

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

## 已支持命令

- `ping`
- `active-document`
- `list-properties`
- `set-property`
- `rebuild`
- `save`

## 示例

```powershell
Invoke-RestMethod http://127.0.0.1:32127/health

Invoke-RestMethod http://127.0.0.1:32127/command `
  -Method Post `
  -ContentType 'application/json' `
  -Body '{"Command":"list-properties"}'
```

## 注册

该项目面向 `.NET Framework 4.8`，用于传统 COM Add-in。

构建后需要用管理员权限注册 DLL：

```powershell
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe ReadBom.SwAddin.dll /codebase
```

取消注册：

```powershell
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe ReadBom.SwAddin.dll /unregister
```
