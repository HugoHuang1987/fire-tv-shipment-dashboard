# Fire TV Shipment Dashboard

本仓库是 Fire TV 出货量数据下载、滚动合并、规则统计、Excel 输出和本地看板的源码包。

仓库只包含程序代码、页面资源和示例配置；不包含原始订单、人员名单、工程师 ID、Excel 结果、日志、token、通知密钥或 LLM key。

## 主要模块

- `tools/Invoke-RdmDownload.ps1`：按人员和时间范围从 RDM 下载原始订单。
- `tools/merge_rdm_snapshot.py`：按订单号滚动合并 RDM 快照，数量变化阻断，机芯按 8 个月窗口内最新快照更新。
- `tools/Invoke-FireTvMonthlyRefresh.ps1`：月度无人值守主流程。
- `tools/Invoke-FireTvRefreshGate.ps1`：每周轻量检查入口，本月已成功时立即退出。
- `tools/Register-FireTvScheduledTask.ps1`：安装每周一运行、错过后补跑的 Windows 计划任务。
- `tools/Update-FireTvTaskStatus.ps1`：读取任务、检查和月度刷新状态，供 Dashboard 展示。
- `dashboard/build_data.py`：按 rule 和 calculation-rules 计算统计明细。
- `dashboard/export_data_source.mjs`：输出统一数据源、异常清单和人工查看 Excel。
- `dashboard/index.html` + `dashboard/assets/`：本地看板页面。
- `dashboard/Sync-FireTvWikiRules.ps1`：Wiki 规则同步审计入口。
- `dashboard/Send-FireTvNotification.ps1`：Windows / ServerChan / 企业微信通知入口。

## 必要环境变量

复制 `config/env.sample.ps1` 为本机配置脚本，按实际路径填写。

关键变量：

- `FIRE_TV_SHARED_ROOT`：共享业务目录，存放 raw、rule、quantity、dashboard、logs 等。
- `FIRE_TV_RDM_BASE_URL`：RDM listing 接口地址。
- `FIRE_TV_WIKI_URL`：Wiki 规则页面地址。
- `FIRE_TV_PYTHON`：Python 可执行文件，可选；未配置时使用 `python`。
- `FIRE_TV_NODE`：Node 可执行文件，可选；未配置时使用 `node`。
- `FIRE_TV_NODE_MODULES`：`@oai/artifact-tool` 所在 node_modules，可选。
- `FIRE_TV_SERVERCHAN_SENDKEY`：个人微信通知 key，可选。
- `FIRE_TV_WECOM_WEBHOOK`：企业微信群机器人 webhook，可选。
- `FIRE_TV_LLM_BASE_URL`、`FIRE_TV_LLM_MODEL`、`FIRE_TV_LLM_API_KEY`：LLM 规则审计接口，可选。

## 典型运行

刷新看板，不跑 Wiki：

```powershell
$env:FIRE_TV_SKIP_WIKI = "1"
.\dashboard\刷新看板数据.cmd "$env:FIRE_TV_SHARED_ROOT"
```

月度刷新：

```powershell
.\tools\Invoke-FireTvMonthlyRefresh.ps1
```

安装 Windows 每周检查任务：

```powershell
.\tools\Register-FireTvScheduledTask.ps1
```

计划任务默认名为 `Fire TV 出货量每周检查`，每周一 09:00 运行，并启用 `StartWhenAvailable`：电脑关机错过计划时间后，会在 Windows 下次可用时补跑。本月已有 `logs/monthly-success-YYYYMM.json` 成功标记时只记录检查，不会重复下载和发布。

Dashboard 右上角的“检查任务状态”会显示最近检测、判断、上次/下次任务时间和最近成功刷新。也可以运行：

```powershell
.\tools\Update-FireTvTaskStatus.ps1 -Show
```

Windows 计划任务直接运行本地 PowerShell 脚本，不需要启动 Codex。

单独测试 RDM 下载：

```powershell
.\tools\Invoke-RdmDownload.ps1 -EngineersFile ".\config\engineers.example.txt" -EngineerIdsCsv ".\config\engineer-ids.example.csv" -StartDate "2026/01/01" -EndDate "2026/07/01" -OnMissing Skip
```

## 不应提交的内容

- 原始订单数据、RDM 快照、补丁数据、统计输出 Excel。
- 人员名单和工程师 ID 表。
- Wiki 快照、冲突清单、告警文件、执行日志。
- Dashboard 数据、刷新状态和 Windows 任务状态快照。
- 任何 token、key、webhook、邮箱或企业内部凭证。
