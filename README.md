# office

Excel 文件处理库。

## xlsx 操作

### 快速开始

```go
// 读取
data, err := xlsx.Read("./test.xlsx")

// 写入
err := xlsx.WriteFile("./new.xlsx", data)
```

### 高级功能

- 原始值/计算值切换
- 字段过滤和映射
- 自定义单元格样式
- 并发处理大数据
- 表头自定义处理

详细文档: [xlsx/README.md](./xlsx/README.md)
