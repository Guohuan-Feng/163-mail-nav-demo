# 163 邮箱基金净值附件提取（演示版）

一个本地、手动执行的 Python 演示工具：从 163 邮箱 INBOX 中读取 .xlsx 附件，提取基金净值数据，并生成固定 10 列的 Excel 台账。

仅用于演示和本地测试，不包含自动发信、计划任务、通知或资金核对功能。

## 功能

- 只读连接 163 邮箱 IMAP，不修改邮件已读状态
- 检索最近指定天数内的 INBOX 邮件
- 下载并解析 .xlsx 附件
- 自动识别常见净值表头及部分同义表头
- 输出固定格式的“净值台账”Excel
- 默认仅输出前 5 条记录，便于现场演示
- 对重复记录去重；同一基金/产品/日期的净值冲突会明确提示
- 支持无需邮箱的离线演示模式

## 环境要求

- Windows + PowerShell
- Python 3.9 或更高版本
- 真实邮箱模式需要可访问 163 邮箱 IMAP 服务

## 安装

~~~powershell
git clone https://github.com/Guohuan-Feng/163-mail-nav-demo.git
cd 163-mail-nav-demo

py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
~~~

如果不使用虚拟环境，也可以直接安装：

~~~powershell
py -m pip install -r requirements.txt
~~~

## 离线演示（推荐先运行）

以下流程不会读取邮箱配置，也不会连接邮箱：

~~~powershell
# 生成两份脱敏的示例 Excel 附件
py nav_mail_demo.py --make-sample-input .\sample_input

# 提取示例附件并生成固定格式台账
py nav_mail_demo.py --input-dir .\sample_input
~~~

输出文件位于：

~~~text
output\基金净值台账_YYYYMMDD_HHMMSS.xlsx
~~~

默认输出恰好 5 条记录。若需输出所有符合条件的记录：

~~~powershell
py nav_mail_demo.py --input-dir .\sample_input --max-rows 0
~~~

仓库中还附有一份脱敏的 [5 行 × 10 列示例输出](./示例输出_5行10列.xlsx)。

## 使用真实 163 邮箱

### 1. 配置授权码

先在 163 邮箱网页设置中开启 IMAP 服务，并生成客户端 IMAP 授权码。

> 使用 IMAP 授权码，不要使用网页登录密码。

复制配置模板：

~~~powershell
Copy-Item config.example.env config.env
notepad config.env
~~~

至少填写以下两项：

~~~text
IMAP_USER=your_163_mail@163.com
IMAP_AUTH_CODE=your_163_imap_authorization_code
~~~

常用配置：

~~~text
LOOKBACK_DAYS=7       # 检索最近邮件天数
MAX_ROWS=5            # 默认最多输出 5 条；设为 0 表示不限制
SUBJECT_KEYWORD=      # 可选：仅处理主题包含该关键词的邮件
DOWNLOAD_DIR=downloads
OUTPUT_DIR=output
MAX_ATTACHMENT_MB=20
~~~

### 2. 执行提取

~~~powershell
py nav_mail_demo.py
~~~

常用参数：

~~~powershell
# 仅检索最近 2 天
py nav_mail_demo.py --days 2

# 最多输出 10 条
py nav_mail_demo.py --max-rows 10

# 输出全部符合条件的记录
py nav_mail_demo.py --max-rows 0

# 指定输出目录
py nav_mail_demo.py --output-dir .\my_output
~~~

## 输出格式

生成的工作簿仅包含一个工作表：“净值台账”。

| 列名 | 说明 |
| --- | --- |
| 基金代码 | 原附件基金代码 |
| 基金名称 | 原附件基金名称 |
| 产品代码 | 原附件产品代码 |
| 产品名称 | 原附件产品名称 |
| 净值日期 | 标准化日期 |
| 单位净值 | 数值 |
| 累计净值 | 数值，可为空 |
| 来源附件 | 原始附件文件名 |
| 来源工作表 | 原始工作表名称 |
| 来源行号 | 原始 Excel 行号，便于追溯 |

## Excel 附件要求与限制

程序会在每个工作表的前 20 行寻找表头。附件至少需要包含：

- 净值日期
- 单位净值
- 基金或产品的代码、名称字段之一

可识别部分常见别名，例如“估值日期”“每份净值”“累计单位净值”。

当前演示版限制：

- 仅支持 .xlsx，会跳过 .xls、.xlsm 等格式
- 默认仅保留前 5 条规范记录
- 对相同“基金/产品/净值日期”的完全重复记录保留首次读取结果
- 如相同主键出现不同净值，程序会提示冲突并保留首次读取结果
- 不包含历史台账合并、应收邮件监控、定时执行或消息通知
- 复杂模板、合并表头、无缓存公式值等 Excel 文件可能需要额外适配

## 安全说明

- config.env 中包含邮箱地址和 IMAP 授权码，**不得上传到 GitHub**
- 仅提交 config.example.env
- 下载的邮件附件和生成的台账可能含敏感数据，不应默认提交
- 程序只读访问 INBOX，不发送邮件、不删除邮件、不创建计划任务
- 测试完成后，建议在 163 邮箱中撤销不再使用的授权码

## 项目结构

~~~text
nav_mail_demo.py      # 主程序
requirements.txt      # Python 依赖
config.example.env    # 配置模板（可安全提交）
README.md             # 使用说明
示例输出_5行10列.xlsx  # 脱敏示例输出
sample_input/         # 离线演示生成的示例附件（运行后生成）
downloads/            # 真实邮箱附件下载目录（运行后生成）
output/               # 台账输出目录（运行后生成）
~~~

## 免责声明

本项目为演示代码。使用真实邮箱及业务数据前，请自行评估数据安全、访问权限和合规要求。
