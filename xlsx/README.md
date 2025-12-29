# xlsx

Excel 文件读写库。

## 功能

- 读取 Excel 文件为 Map 结构
- 写入 Map 数据到 Excel 文件
- 支持原始值、公式值、计算值
- 并发处理大数据
- 自定义单元格样式
- 字段过滤和映射

## 读取

### 基本用法

```go
// 读取整个文件
data, err := xlsx.Read("./test.xlsx")

// 读取并处理每一行
data, err := xlsx.Read("./test.xlsx", func(opt *xlsx.ReadOptions) {
    opt.Handler = func(row int, data ztype.Map) ztype.Map {
        // 自定义处理逻辑
        return data
    }
})
```

### 读取选项

| 选项 | 类型 | 说明 |
|------|------|------|
| Sheet | string | 工作表名称，默认第一个 |
| Fields | []string | 只读取指定字段 |
| RawCellValueFields | []string | 指定字段使用原始值（公式） |
| CalcCellValueFields | []string | 指定字段使用计算值 |
| RawCellValue | bool | 全部使用原始值 |
| HeaderHandler | func | 自定义表头处理 |
| HeaderMaps | map[string]string | 表头映射 |
| Reverse | bool | 反向读取 |
| Parallel | uint | 并发数，0=自动 |
| OffsetX, OffsetY | int | 跳过前 N 列/行 |
| NoHeaderRow | bool | 无表头模式 |
| MaxRows | int | 最大读取行数 |
| RemoveEmptyRow | bool | 移除空行 |
| TrimSpace | bool | 去除首尾空格 |

### 示例

```go
// 读取指定字段
data, err := xlsx.Read("./test.xlsx", func(opt *xlsx.ReadOptions) {
    opt.Fields = []string{"name", "age"}
    opt.MaxRows = 100
})

// 表头映射
data, err := xlsx.Read("./test.xlsx", func(opt *xlsx.ReadOptions) {
    opt.HeaderMaps = map[string]string{
        "姓名": "name",
        "年龄": "age",
    }
})

// 获取公式而非计算值
data, err := xlsx.Read("./test.xlsx", func(opt *xlsx.ReadOptions) {
    opt.RawCellValueFields = []string{"total", "sum"}
})
```

## 写入

### 基本用法

```go
data := ztype.Maps{
    {"name": "张三", "age": 25},
    {"name": "李四", "age": 30},
}

// 写入文件
err := xlsx.WriteFile("./output.xlsx", data)

// 写入字节
buf, err := xlsx.Write(data)
```

### 写入选项

| 选项 | 类型 | 说明 |
|------|------|------|
| Sheet | string | 工作表名称，默认 Sheet1 |
| First | []string | 首列字段优先 |
| Last | []string | 末列字段优先 |
| CellHandler | func | 自定义单元格样式 |

### 示例

```go
// 控制列顺序
err := xlsx.WriteFile("./output.xlsx", data, func(opt *xlsx.WriteOptions) {
    opt.First = []string{"id", "name"}  // 前两列
    opt.Last = []string{"created_at"}   // 最后一列
})

// 自定义样式
err := xlsx.WriteFile("./output.xlsx", data, func(opt *xlsx.WriteOptions) {
    opt.CellHandler = func(sheet, cell string, value interface{}) ([]xlsx.RichText, int) {
        // 返回富文本和样式ID
        return nil, styleID
    }
})
```

## 高级用法

### 使用句柄

```go
// 打开文件
f, err := xlsx.Open("./test.xlsx")
if err != nil {
    panic(err)
}
defer f.Close()

// 读取
data, err := f.Read()

// 写入
err = f.WriteFile("./new.xlsx", data)

// 访问底层引擎
engine := f.Engine()
// 使用 excelize 原生功能
```

### 创建样式

```go
f, _ := xlsx.Open("./test.xlsx")
defer f.Close()

styleID, err := f.NewStyle(&excelize.Style{
    Font: &excelize.Font{Bold: true},
    Fill: excelize.Fill{Type: "pattern", Color: []string{"#E0E0E0"}},
})
```
