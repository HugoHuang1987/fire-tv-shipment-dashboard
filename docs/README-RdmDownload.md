# RDM 下载器

## 一键运行

双击：

```bat
tools\Run-RdmDownload-SelectedEngineers.cmd
```

默认读取：

- 人员名单：`tools\rdm_engineers.txt`
- 人员 ID：`tools\rdm_engineer_ids.csv`
- 日期范围：`2023/10/01` 到今天
- 输出目录：`outputs\rdm_download_yyyyMMdd_HHmmss`

输出文件：

- `data.csv`：兼容旧下载器格式的原始 CSV
- `summary.csv`：每个人下载了多少行、是否成功
- `raw_rows.jsonl`：原始接口行数据，便于以后排查
- `missing_engineer_ids.csv`：缺 ID 的人员
- `urls.txt`：请求过的接口 URL
- `run.log`：运行日志

## 常用命令

指定人员：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-RdmDownload.ps1 -Engineers "黄海明 赵龙" -PageSize 200
```

指定日期：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-RdmDownload.ps1 -StartDate "2023/10/01" -EndDate "2026/06/25"
```

只检查 URL，不真正下载：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-RdmDownload.ps1 -DryRun
```

临时补一个人员 ID：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-RdmDownload.ps1 -Engineers "赖锦灵:1234"
```

如果公司网络必须走系统代理：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-RdmDownload.ps1 -UseSystemProxy
```

使用源码内置的完整人员 ID 表：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-RdmDownload.ps1 -EngineerIdsCsv .\tools\rdm_engineer_ids_source.csv -AllFromIds
```
