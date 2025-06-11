# office

## xlsx 操作

```go
// 读取
data, err := xlsx.Read("./test.xlsx")

// 写入
err := xlsx.WriteFile("./new.xlsx", data)
```
