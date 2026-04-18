# 后端服务（可选）

前端 index.html 已包含完整功能，**无需启动后端即可使用**。

后端适用于以下场景：
- 需要服务端批量处理大量 Excel 文件
- 需要精确控制 Excel 打印样式（openpyxl 比 SheetJS 支持更多 Excel 特性）
- 需要集成到现有 Flask 服务中

## 启动



服务地址：http://127.0.0.1:5188

## API

| 接口 | 方法 | 说明 |
|------|------|------|
| /api/health | GET | 健康检查 |
| /api/analyze | POST | 上传 Excel，返回统计信息 |
| /api/process | POST | 上传 Excel，下载带分页符的文件 |

所有 POST 接口使用 multipart/form-data，文件字段名为 file。
