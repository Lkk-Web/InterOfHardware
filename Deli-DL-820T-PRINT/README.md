# 德力 DL-820T 标签打印服务

## 简介

德力 DL-820T 标签打印服务是一个专为德力 DL-820T 标签打印机设计的打印服务程序。该服务提供了一个 HTTP API 接口，允许其他应用程序通过网络发送打印请求，实现批量打印标签的功能。

## 功能特点

- 支持批量打印标签
- 自动生成包含批次、包号等信息的二维码
- 提供 RESTful API 接口，方便集成到其他系统
- 自动检测打印机连接状态
- 支持开机自启动

## 设置开机自启动

- 将`invisible_start.vbs`文件的`快捷方式`放入 Windows 启动文件夹,通过按下`Win+R`，然后输入`shell:startup`快速访问此文件夹。

- **BAT 文件(.bat)**: 批处理文件是 Windows 系统中的脚本文件，包含一系列命令，可以自动执行重复性任务。在本项目中，BAT 文件用于检查环境依赖并启动 Python 打印服务。

- **VBS 文件(.vbs)**: Visual Basic Script 文件是 Windows 系统中的脚本文件，可以执行更复杂的操作。在本项目中，VBS 文件用于在后台无窗口模式下启动 BAT 文件，使服务启动过程`对用户完全透明`。

## 测试用 JSON 数据

以下是一个用于测试标签打印 API 的 JSON 数据示例：

```json
{
  "dto": {
    "batch": "WGC2025031801",
    "product_name": "Z0118",
    "manufacturer": "RY",
    "date": "2025-3-17",
    "license_plate": "粤A12345"
  },
  "start_package": 100,
  "end_package": 102
}
```

您可以使用 Postman 或 curl 等工具发送这个 JSON 到您的 API 端点：

```
POST http://localhost:5010/print
Content-Type: application/json
```

这个请求将会打印 3 个标签，包号从 100 到 102。如果您想测试更多或更少的标签，只需调整`start_package`和`end_package`的值即可。
