# 瑞典 NPA 管理员自动化工具

这个工具按《瑞典包装服务商后台新增管理员 SOP》执行客户账号维护：登录 NPA 门户，读取邮箱验证码，检查并添加管理员，检查并修改 Invoice email。

当前版本：`v2026.09.22`。验证码只走自动读取与提交流程，图形界面不再提供人工输入和“提交验证码”按钮。

## EXE 版使用方法

1. 解压 `RobotAdmin_single_release_20260920.zip`。
2. 确认电脑已经安装 Microsoft Edge。
3. 双击 `启动机器人.bat`。它会启动同目录下的新版 `RobotAdmin.exe`；后台执行组件和默认配置已经嵌入主程序，不需要另找或手动启动 `robot_worker.exe`。
4. 选择客户 `.xlsx` 表格，按需要选择是否显示浏览器窗口。
5. 点击“开始运行”，执行正式流程。

电脑不需要安装 Python。轻量版使用已安装的 Microsoft Edge，不再内置 Chromium。

## 自动执行内容

每个客户按以下顺序处理：

1. 打开 NPA 门户并输入客户邮箱。
2. 登录对应邮箱入口，按登录时间筛选最新的 NPA 验证码邮件。
3. 读取邮件正文中的账户验证码并提交到 NPA。
4. 打开 Settings（瑞典语界面为 `Inställningar`）。
5. 进入 Users（瑞典语界面为 `Användare`）。
6. 检查 `admin@example.com` 是否已经存在；不存在时添加 `Target Admin`，权限选择 Administrator。
7. 读取 Invoice email；如果不是 `invoice@example.com`，修改并保存。
8. 对 Add User 和 Invoice email 保存结果进行回读确认。

已存在的管理员不会重复添加，已经正确的 Invoice email 不会重复修改。

对于页面加载、邮箱状态或验证码提交等运行时故障，单个客户最多执行 3 次完整流程（首次尝试加 2 次重试）。每次重试都会重新创建浏览器上下文，避免上一次的页面状态影响下一次。第 1 或第 2 次成功时不会进入失败表；连续 3 次仍失败时，才会写入 `failed.xlsx`，并在 `failure_reason` 中保留每次失败的步骤和原因。输入缺失、邮箱入口无法匹配等确定性问题不会重复重试。

## 客户表格

GUI 支持 `.xlsx` 和 `.csv`。推荐使用以下字段：

| 字段 | 必填 | 说明 |
|---|---:|---|
| `customer_name` | 否 | 客户名称，用于日志和结果表识别。 |
| `portal_email` | 是 | NPA 门户登录邮箱。也支持中文列名 `邮箱`。 |
| `portal_password` | 否 | NPA 密码；验证码登录时可以为空。 |
| `mail_email` | 是 | 接收验证码的邮箱账号；为空时使用 `portal_email`。 |
| `mail_password` | 是 | 邮箱密码；为空时使用 `portal_password`。 |
| `enabled` | 否 | 填 `false`、`否` 或 `跳过` 时跳过该行。 |

当前表格中的中文列名也支持：`客户中文名称`、`客户英文名称`、`邮箱`、`邮箱密码`。

## 支持的邮箱入口

工具会根据邮箱域名自动选择入口：

- `erp-aid.example.invalid`：`mail.erp-aid.example.invalid`
- `erp-helper.example.invalid`、`example.invalid`：`mail.erp-helper.example.invalid`
- `eu-erp.example.invalid`：`mail.eu-erp.example.invalid`
- `ecopv-info.example.invalid`：`mail.ecopv-info.example.invalid`
- `163.com`：网易邮箱
- `ecopv0316.example.invalid`、`mail2.example.invalid`、`mail3.example.invalid`：阿里企业邮箱

## 运行结果

结果会写入程序目录下的 `runs\日期_时间\`，每处理完一条客户就立即写入，不等整批结束：

- `logs\operator.log`：面向操作人员的简洁进度日志，包含当前客户、当前步骤和成功/失败/跳过统计。
- `logs\events.jsonl`：详细事件日志，不记录密码和验证码。
- `results\results.xlsx`：本批全部输入记录，未处理的行会明确标为 `pending/未处理`；每完成一条就立即写入状态、尝试次数、管理员动作、Invoice email 动作、失败步骤、失败原因和更新时间。
- `failed\failed.xlsx`：只包含连续 3 次尝试仍失败或需要人工处理的记录；每条都带有 `status_label`、`attempts`、`failed_step`、`failure_reason`、`retryable`，方便人工处理或单独重跑。管理员已存在的记录不会进入失败表。
- `state\run_summary.json`：本次运行的实时汇总和最后处理位置。

结果表和原始客户表格会同步写入以下结果列：`row_number`、`status`、`status_label`、`attempts`、`admin_action`、`admin_result`、`invoice_action`、`invoice_result`、`failed_step`、`failure_reason`、`error`、`retryable`、`updated_at`、`run_id`。程序不会另存原表备份；写入时只使用同目录临时文件完成替换，避免中途写入半个文件。

如果操作人员中途停止或浏览器异常，已完成客户的结果仍保留在上述目录和原表中；未出现结果的客户不会被伪造为成功。

结果状态：

- `completed / 已完成`：客户流程完成。
- `pending / 未处理`：尚未轮到该记录，或运行在此处中断。
- `manual_required / 待人工处理`：缺少必要输入、入口无法匹配或需要人工确认。
- `failed / 失败`：该客户在某个步骤失败，需要查看 `failed_step` 和 `failure_reason`。

管理员动作和发票邮箱动作会分别写入代码列与中文说明列：

- `already_exists / 已存在，跳过添加`：目标管理员已经存在，不会重复添加。
- `added / 已添加并确认`：Add User 后已在 Users 列表回读到目标管理员。
- `already_correct / 原值正确，无需修改`：Invoice email 已经是目标地址。
- `updated / 已修改并确认`：Invoice email 修改后已回读确认。

## 常见问题

### 浏览器没有启动

确认使用的是新包中的 `RobotAdmin.exe`。后台组件已经嵌入，不需要旁边再放 `robot_worker.exe` 或 `config.json`。Edge 必须已经安装。

### 自动读取验证码失败

程序会重新创建页面并最多重试 3 次；只有 3 次都失败才会将该客户写入失败表。请在 `failed.xlsx` 的 `failed_step` 和 `failure_reason` 中查看具体卡点。

### 某个客户失败，其他客户正常

工具会记录该客户失败并继续处理后续客户。修复网络或页面问题后，可以再次运行；已经存在的管理员和正确的 Invoice email 会自动跳过。

### 结果中显示 `admin_exists_skip`

这不是失败，表示管理员已经存在，程序为了避免重复添加而跳过 Add User。

### 结果中显示 `already_correct`

这不是失败，表示 Invoice email 已经是 `invoice@example.com`，无需再次保存。

## 开发运行

如果需要从源码运行：

1. 安装 Python 3.11 或更高版本。
2. 安装 `requirements.txt` 中的依赖。
3. 使用 `start_robot.bat` 启动 GUI。

源码版默认使用本机 Playwright 浏览器；发布版使用电脑已安装的 Microsoft Edge。发布版不需要 Python，也不内置 Chromium。
