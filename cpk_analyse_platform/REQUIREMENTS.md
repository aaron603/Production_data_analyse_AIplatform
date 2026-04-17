# 产线数据分析AI平台 — 需求说明文档 v2.2

**项目**：Zillnk Efficiency Improvement Group — 产线数据分析AI平台  
**代码目录**：`D:\Gitlab\cpk_analyse_platform`  
**更新日期**：2026-04-17

---

## 一、产品概述

面向Zillnk质量工程师的桌面工具（Python tkinter GUI），覆盖三大数据来源：

| 功能 | 数据来源 | 状态 |
|------|---------|------|
| 功能一：本地测试站数据分析 | 本地站测试数据（xlsx/json/log） | **已实现** |
| 功能二：深科技 MES 数据分析 | 深科技 MES 导出数据 | 待实现（规划中） |
| 功能三：立讯 MES 数据分析 | 立讯 MES 导出数据 | 待实现（规划中） |

顶层窗口为 `ttk.Notebook`，三个 Tab 对应三大功能：

```
[ 本地数据分析 ] [ 深科技MES分析 ] [ 立讯MES分析 ]
```

---

## 二、GUI 布局（功能一）

### 2.1 整体布局（从上到下）

```
┌─────────────────────────────────────────────────────┐
│  Section 1：输入 / 输出配置                          │
├──────────────────────────────┬──────────────────────┤
│  Section 2：测试工站配置      │  Section 3：分析模式  │
│  （含工站合并，默认折叠）     │  （含文件类型选择）   │
├──────────────────────────────┴──────────────────────┤
│  Section 4：操作 + 进度条                            │
└─────────────────────────────────────────────────────┘
运行日志：独立弹出窗口（点击"查看运行日志"按钮打开）
```

> **关键设计**：
> - Section 2 与 Section 3 **左右并列**显示（左侧宽度弹性，右侧固定 270px）
> - 窗口高度**自适应内容**（启动时自动计算所需高度，不留空白）
> - 运行日志为**独立 Toplevel 弹窗**（760×420，深色主题，关闭后隐藏不销毁）

---

### 2.2 Section 1：输入 / 输出配置

两个字段并排显示在同一行，输入框随窗口宽度自动伸缩。

| 字段 | 说明 |
|------|------|
| DUT条码（可选） | "最后一次pass数据"模式必填；"所选文件夹分析"遍历模式若填写则只分析Excel中的条码；其余模式忽略。第一列=序列号（主条码），第二列=产品编码（可选）。 |
| 输出目录 | 分析结果存放根目录。每次分析自动在此目录下创建 `{产品类别}_{YYYYMMDD_HHMMSS}` 子文件夹，本次所有输出均存入其中，互不覆盖。产品类别取工站文件夹下 `TestResult/` 的第一级非调试子目录名称；若无法识别弹框提示并退化为纯时间戳命名。 |

---

### 2.3 Section 2：测试工站配置

#### 2.3.1 工站列表

每行：`[工站类型 Entry(宽10)] [测试数据文件夹 Entry(可扩展)] [浏览] [删除]`  
底部：`[＋ 添加工站]` 按钮  
支持 ↑/↓ 键在行间切换焦点（同列对齐）。

#### 2.3.2 工站合并配置（默认折叠）

**需求背景**：同一类产品测试时可能有多台同类工站（FT1_1、FT1_2…），若测试内容完全相同需合并分析。

**入口**：Section 2 底部分隔线下方，点击 `▶ 工站合并配置` 展开。

```
展开后显示：
  说明：将指定工站类型的数据合并到目标工站类型中分析
  ┌──────────┐  → 合并到  ┌──────────┐  [✕]
  │  FT2     │           │  FT1     │
  └──────────┘           └──────────┘
  [＋ 添加合并规则]
```

**逻辑规则**：
- 每条规则：`源工站类型 → 目标工站类型`
- 合并后，源工站的数据在提取和分析阶段归入目标工站类型
- 支持多条规则（如 FT2→FT1、FT3→FT1）
- 默认无规则（功能默认关闭）
- 配置持久化到 `app_config.json`

---

### 2.4 Section 3：分析模式

#### 2.4.1 分析文件类型（单选）

```
分析文件类型：  ● xlsx   ○ json
```

- 决定 CPK 分析时读取 xlsx 还是 json 子文件夹
- 默认选 xlsx，配置持久化
- 同一次分析中 xlsx 与 json 文件数量可能不同（部分产品无 json），属正常现象

#### 2.4.2 分析模式下拉（5 种）

| 模式显示名 | 内部值 | DUT条码Excel | 触发故障分析 |
|-----------|--------|------------|------------|
| 最后一次pass数据 | `latest_pass` | **必填** | 否 |
| 所选文件夹分析 | `folder_direct` | 可选（用于条码过滤） | 否（含失败时生成失败分析报告） |
| 全部成功数据 | `all_pass` | 否 | 否 |
| 所有数据（含失败） | `all_with_fail` | 否 | **是** |
| 仅失败数据 | `fail_only` | 否 | **是** |

鼠标悬停下拉框显示对应模式详细说明 Tooltip。

#### 2.4.3 故障分析方式（仅 all_with_fail / fail_only 时显示）

```
故障分析方式：  [基础版（规则库）  ▼]
```

| 选项 | 说明 |
|------|------|
| 基础版（规则库） | 仅使用 fault_rules 表中的关键词规则匹配 |
| 增强版（规则库+Ollama） | 规则库 + 本地 Ollama LLM 辅助分类（需安装 Ollama） |

---

### 2.5 Section 4：操作

- `[开始分析]` 按钮（运行中变为 `[停止分析]`，可随时中止）
- `[查看运行日志]` 按钮（打开/激活日志弹窗）
- 进度标签（显示当前步骤和百分比）
- 进度条（0–100%）

#### 运行日志弹窗

- 独立 Toplevel 窗口，默认隐藏，点按钮显示
- 深色主题（bg `#1e1e2e`，fg `#a8d8a8`）
- 内置"清空"按钮
- 关闭弹窗 = 隐藏（不销毁），再次打开自动滚到最新行
- 同时写入 `analysis_log_<时间戳>.txt` 文件

#### 异常弹框

- 读取发货 Excel 失败 → `messagebox.showerror` 提示文件占用/格式问题
- `_run_analysis` 主流程任意异常 → 弹框显示异常信息，引导查看运行日志

---

### 2.6 菜单栏

| 菜单 | 子项 | 功能 |
|------|------|------|
| 工具 | 加载故障关系描述文件… | 导入 YAML 格式的故障规则文件到 fault_database.db（支持增量合并） |
| 帮助 | 使用帮助 | 弹出使用说明文本窗口 |
| 帮助 | 关于 | 版本信息 |

---

## 三、分析模式详细说明

### 模式 1：最后一次pass数据（`latest_pass`）

**触发条件**：需 DUT条码Excel（第一列序列号、第二列产品编码）

**处理流程**：
```
1. 读取 DUT条码Excel → 获得条码列表
2. 遍历所有配置工站的 TestResult 目录结构
3. 对每个条码，找到所有时间戳目录，筛选最新一次全pass记录
   （时间优先取文件名中14位时间戳，比文件内start_time更可靠）
4. 提取 xlsx 和 json 文件（保持原文件名）
5. 按工站类型分类存入本次运行子目录：
     {run_dir}/{station_type}/xlsx/{原文件名.xlsx}
     {run_dir}/{station_type}/json/{原文件名.json}
   （若合并规则有效，源站数据写入目标站目录）
6. 生成 missing_barcodes.xlsx（未找到pass记录的条码清单）
7. 对每类工站下的 xlsx/json 文件夹执行 CPK 分析
8. 生成 HTML 报告
```

---

### 模式 2：所选文件夹分析（`folder_direct`）

**触发条件**：无需 Excel（可选填 DUT条码Excel 用于条码过滤）

**处理流程**：
```
1. 若填写了 DUT条码Excel，只处理 Excel 中列出的条码；否则处理全部条码
2. 遍历所有配置工站目录（pass + fail 全量提取）
   调用 run_extraction_traverse()，使用 _walk_all_records_in_folder() 递归扫描
   每个工站文件夹下 barcode/timestamp 结构，收集所有记录
3. 对提取出的文件执行 CPK 分析
4. 生成 cpk_report.html + comprehensive_report.html（传入 fail_data）
5. 若存在失败记录，额外生成：
     folder_direct_fail_analysis.xlsx（3 Sheet：失败条码 / 失败测试项 / 从未成功条码）
     fail_analysis_report.html（帕累托图 + 汇总卡片 + 失败条码表 + 从未成功条码表）
```

**适用场景**：测试数据存储在标准 TestResult 目录结构中，需提取后分析。

> **注意**：此模式始终执行多层目录遍历（`run_extraction_traverse`），不再区分"文件夹直接含目标文件"与"TestResult结构"两种场景。

---

### 模式 3：全部成功数据（`all_pass`）

**触发条件**：无需 Excel

**处理流程**：
```
1. 直接遍历所有配置工站目录，无需发货Excel
2. 对每个条码的每个时间戳目录，只要该次测试全部通过就提取
   （一个条码可能被提取多次，对应多次成功测试记录）
3. 提取 xlsx/json 到输出目录（同模式1的分类方式）
4. 生成 duplicate_barcodes.xlsx（重复测试条码统计，列出各条码的测试次数和各次时间）
5. 对每类工站做 CPK 分析，生成 HTML 报告
```

**与模式1的区别**：模式1每个条码只取最新一次pass；模式3取所有pass记录，体现完整过程能力。

**额外输出**：`duplicate_barcodes.xlsx`（分工站类型，含：条码、工站类型、重复次数、各次测试时间；按重复次数降序）

---

### 模式 4：所有数据（含失败）（`all_with_fail`）

**触发条件**：无需 Excel；**自动触发故障分析**

**核心价值**：跨站比对 — 同一模块在不同设备的测试数据关联分析

**处理流程**：
```
1. 自动发现全部条码，遍历所有配置工站目录，收集全部记录（pass + fail）
2. Pass 记录写入故障库（fault_type='测试通过'），供跨站比对参考
3. Fail 记录完整提取：failed_items、equip_errors、first_fail_desc 等
4. 通过 station_machine（EQP_ID）识别跨站记录，分析站级/DUT级问题
5. 规则匹配（+ 可选 Ollama LLM），写入 fault_database.db
6. 生成 fault_barcodes.xlsx（故障条码列表，含故障类型/失败测试项数等）
7. 生成 rule_suggestions_*.yaml（未分类故障 + 高频失败项规则建议模板）
8. 生成包含失败数据的 CPK HTML 报告（蓝pass/红fail直方图）
```

**跨站推断逻辑**：

| 观察到的跨站模式 | 推断 |
|----------------|------|
| BC001 在 FT1_1 失败（COM6错误），在 FT1_2 通过 | **站级问题**：FT1_1 设备故障 |
| BC001 在 FT1_1 失败（RX偏低），在 FT1_2 也失败（同项） | **DUT问题**：产品本身硬件缺陷 |

**额外输出**：`fault_barcodes.xlsx` + `rule_suggestions_<时间戳>.yaml`

---

### 模式 5：仅失败数据（`fail_only`）

**触发条件**：无需 Excel；**自动触发故障分析**

**核心价值**：快速总结失败规律，完善规则库

**处理流程**：
```
1. 快速预筛选：只处理 fail 记录（跳过所有 pass，速度更快）
2. 对每条 fail 记录提取结构化故障数据
3. 规则匹配 + 可选 LLM 分析，输出 fault_type
4. 生成故障模式统计：
     - 高频失败测试项 TOP20
     - 高频设备/通信错误 TOP10
     - 故障分类分布
     - 未分类故障样本（最多20条）
5. 生成 fault_barcodes.xlsx（故障条码列表）
6. 生成 rule_suggestions_*.yaml（规则建议模板）
```

**额外输出**：`fault_barcodes.xlsx` + `rule_suggestions_<时间戳>.yaml`

> **设计意图**：先用 `fail_only` 快速总结规律、完善规则库；再用 `all_with_fail` 建立含跨站比对的完整故障分析库。

---

## 四、测试数据目录结构

### 4.1 生产标准目录层级

```
{station_root}/                      ← 用户在UI配置的工站文件夹
  TestResult/                        ← 固定目录名
    {product_category}/              ← 产品类别（ORBI_B3、ORBI_B40等）
      {station_type}/                ← 工站类型（FT1、FT2、Aging、VSWR等）
        debug*/                      ← 自动跳过（调试目录）
        {product_code}/              ← 产品编号/版本（X11_X11、R1B等）
          {barcode}/                 ← 单条码 或 双条码(BC1_BC2)
            {YYYYMMDDHHMMSS[_sfx]}/  ← 单次测试记录（时间戳目录）
              Test_Result_*_{BC}.xlsx     ← 主测试数据
              *_MEASUREMENT_Zillnk.json  ← B3B40/PA板，含首次失败描述
              ate_test_log.log            ← 主日志
              ate_test_log.html
              Failed_points_*.txt         ← RRU类，失败项摘要
              file_bk/
                env_config.yml            ← 仪器VISA地址、EQP_ID
                env_comp/*.csv            ← 链路插损校准数据（freq,s21）
                test_limits.yml           ← 测试限值配置
                test_cases.yml            ← 测试用例配置
              RU1_Log_{BC}/               ← DUT通信日志
              TM1_Log/                    ← 仪器通信日志
              {TestItemName}/             ← 截图目录（以测试项命名）
              screen_*.png
              DutInfo_Mes=*.json
```

### 4.2 目录遍历策略

| 场景 | 处理方式 |
|------|---------|
| 双条码目录 `BC1_BC2` | 取第一个条码为主条码，`barcode_full` 保存完整名称 |
| 时间戳带后缀 `20251008163411_NT` | `_NT` 等后缀为测试类型标识，解析时忽略后缀 |
| 调试目录 `debug/Debug/DEBUG` | 自动跳过，不计入分析 |
| 无 TestResult 子目录 | 降级为直接扫描条码目录（兼容拷贝数据） |
| 工站合并规则有效 | 遍历时将源站数据归入目标站类型 |
| 多台并行测试站时间戳 | **优先使用文件名中的14位时间戳**（比文件内start_time更可靠，避免并行站时序错乱） |

### 4.3 xlsx 与 json 格式说明

`Test_Result_*_{BC}.xlsx`：
- 每个 sheet 页对应一个测试项（sheet名 = 测试项名）
- 列：`point_name`（测试子项）、`data`（测量值）、`limit_high`(USL)、`limit_low`(LSL)、`result`、`start_time`、`stop_time`、`station`

`*_MEASUREMENT_Zillnk.json`：
- 结构：`DutInfo` + `TestResult[].CaseName` + `TestResult[].TestPoints[]`
- `DutInfo` 字段含：`ProductName`、`SiteName`（EQP_ID）、`Result`、`FirstFailCaseDescription`
- `TestPoints` 字段：`TestPointNumber`、`TestData`、`LimitLow`、`LimitHigh`、`Result`

---

## 五、输出文件

每次点击"开始分析"，在用户配置的输出目录下自动创建 **`{产品类别}_{YYYYMMDD_HHMMSS}/`** 子目录，本次所有输出均存入其中，历次运行互不覆盖。

- **产品类别**：取各工站文件夹下 `TestResult/` 的第一级非调试子目录名（如 `ORBI_B3`）
- 若无法识别 TestResult 结构，弹框提示并退化为纯时间戳命名

```
{out_dir}/
  {产品类别}_{YYYYMMDD_HHMMSS}/              ← 本次运行专属子目录
    {station_type}/xlsx/                      提取的测试 .xlsx 文件（原文件名）
    {station_type}/json/                      提取的测试 .json 文件（原文件名）
    missing_barcodes.xlsx                     缺失/异常条码汇总（latest_pass/all_with_fail/fail_only）
    duplicate_barcodes.xlsx                   重复测试条码统计（all_pass 模式）
    fault_barcodes.xlsx                       故障条码列表（all_with_fail/fail_only 模式）
    folder_direct_fail_analysis.xlsx          失败条码明细（folder_direct 场景B）
    rule_suggestions_<时间戳>.yaml            规则建议模板（all_with_fail/fail_only 模式）
    cpk_report.html                           CPK 专项报告（按工站+测试大项分组）
    comprehensive_report.html                 综合分析报告（6 Tab：总览/失败分析/CPK/数据分布/失败模式/故障回放；所有模式）
    fail_analysis_report.html                 失败分析报告（folder_direct 场景B，含失败时）
    fault_database.db                         故障分析关系库（SQLite，跨次分析持续积累）
    analysis_log_<时间戳>.txt                 本次运行完整过程日志
```

### fault_barcodes.xlsx 列说明

| 列名 | 说明 |
|------|------|
| 条码 | 主条码 |
| 完整条码 | 原始目录名（含双条码） |
| 工站 | 用户配置的工站标签 |
| 机台 | 物理设备编号（EQP_ID） |
| 测试时间 | 记录时间戳 |
| 状态 | fail / unknown |
| 故障类型 | 规则匹配/LLM分析结果 |
| 失败测试项数量 | failed_items 条数 |
| 首次失败描述 | ATE FirstFailCaseDescription |

按测试时间倒序排列。

### rule_suggestions_<时间戳>.yaml 结构

```yaml
# 系统自动生成的规则建议模板
rules:
  # 未分类故障样本（最需要填写）
  - keywords: "RX NF ORBI 3600"    # 出现 N 次
    fault_type: ""    # TODO: 请填写故障类型
    suggestion: ""    # TODO: 请填写处置建议

  # 高频失败测试项
  - keywords: "PA IDLE CURR"        # 出现 N 次
    fault_type: ""
    suggestion: ""

  # 设备/通信错误参考（注释行，仅供参考）
  # - keywords: "串口设备断连"       # 出现 N 次
```

---

## 六、HTML 报告（已实现）

每次分析结束后，程序在输出目录生成两份报告，并自动在浏览器打开综合报告。

### 6.1 综合报告（`comprehensive_report.html`）

由 `core/html_comprehensive_report.py` → `generate_comprehensive_report()` 生成，使用 Chart.js（需浏览器访问 cdn.jsdelivr.net）。

**6 个 Tab 内容**：

| Tab | 内容 |
|-----|------|
| 总览 | KPI卡片（总样本量/整体良率/失败条码数/CPK达标率）、良率趋势折线、失败类型饼图、测试大类汇总表 |
| 失败分析 | Top 25 高频失败测试项柱图（点击可查看明细）、失败记录明细表（含SN搜索） |
| CPK分析 | Cpk 横向柱图（色标分级 <1.0红/1.0-1.33橙/≥1.33绿）、完整 CPK 统计表（含搜索） |
| 数据分布 | 按测试项切换的堆叠直方图（pass蓝/fail红）、统计面板（n/均值/σ/LSL/USL） |
| 失败模式 | 失败类型统计卡片、按小时热图（时序分布）、多失败项SN分析 |
| 故障回放 | 左侧SN列表（可搜索/过滤状态）、右侧逐Sheet展开测试结果 |

- `fail_data` 参数有值时（`folder_direct` 模式），图表含 fail 着色；否则全 pass 数据
- 自动在分析完成后 0.8 秒用默认浏览器打开

### 6.2 CPK 专项报告（`cpk_report.html`）

由 `core/html_report.py` → `generate_report()` 生成，自包含无外部依赖。

按工站类型 + 测试大项分组展示 CPK 统计表和分布图。

### 6.3 失败分析报告（`fail_analysis_report.html`）

由 `core/html_fail_report.py` → `generate_fail_report()` 生成，仅在 `folder_direct` 模式且存在失败记录时生成。

内容：帕累托图 + KPI 汇总卡片 + 失败条码表 + 从未成功条码表。

### 6.4 ProductName 获取优先级

1. `TestResult/` 下一级目录名（如 `ORBI_B3`）
2. 各工站文件夹路径中的路径段
3. 兜底：空字符串（仅用时间戳命名输出目录）

### 6.5 待实现的报告增强（P1 规划）

- 两级 Tab 结构：第一级工站类型 → 第二级 xlsx sheet 名
- CPK 列默认隐藏，搜索框输入 `cpk` 后显示
- 产品数据检索 Tab（三层展开：条码→测试项→子项明细）
- 范围搜索：指定测试子项 + 数值范围，查询命中条码

---

## 七、CPK 分析

### 7.1 数据来源

- 按分析文件类型选择（xlsx 或 json）读取提取目录下的文件
- 每类工站独立分析（合并规则有效时，合并后视为同一工站）
- **LSL/USL 取最新测试文件中的限值**（按文件名14位时间戳排序，不依赖文件内时间）

### 7.2 子项过滤（跳过不做 CPK）

- 结果为字符串型（非数值）
- 所有值都相同的固定值项（std == 0）
- 样本数不足（< 2）

### 7.3 统计量

| 类别 | 字段 |
|------|------|
| 基础（始终显示） | 样本数(n)、均值(μ)、标准差(σ)、min、max、LSL、USL、通过率、通过数/失败数 |
| CPK（默认隐藏） | Cp、Cpl（下单边）、Cpu（上单边）、Cpk |

### 7.4 values 三元组格式

CPK 计算结果中，`values` 字段存储 `(barcode, measurement_value, is_pass)` 三元组，用于直方图着色（蓝/红）和通过率统计。

---

## 八、故障分析模块

### 8.1 数据库表结构（SQLite）

**`fault_rules`**（规则表，支持 CRUD + 用户文件导入）：

| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER | 自增主键 |
| keywords | TEXT | 逗号分隔关键词 |
| fault_type | TEXT | 故障分类标签 |
| suggestion | TEXT | 建议处置措施 |
| created_at / updated_at | TEXT | 时间戳 |

**`fault_records`**（每条测试记录一行）：

| 字段 | 类型 | 说明 |
|------|------|------|
| barcode | TEXT | 主条码 |
| barcode_full | TEXT | 完整条码（含双条码） |
| station | TEXT | 用户配置的工站标签 |
| station_machine | TEXT | 物理设备编号（来自 EQP_ID） |
| product_category | TEXT | 产品类别（ORBI_B3 等） |
| product_code | TEXT | 产品编号（X11_X11 等） |
| test_time | TEXT | 测试时间 |
| status | TEXT | pass / fail / unknown |
| fault_type | TEXT | 故障分类 |
| first_fail_desc | TEXT | ATE首次失败描述（MEASUREMENT JSON） |
| failed_items | JSON | 结构化失败项（含偏差方向和量） |
| equip_errors | JSON | 检测到的设备/通信错误 |
| instruments | JSON | 仪器配置（VISA地址映射） |
| log_excerpt | TEXT | 日志关键行摘要 |
| llm_analysis | JSON | Ollama增强分析结果 |

**`fault_stats`**：各类故障计数汇总

### 8.2 分析流程

```
ate_test_log.log
  ↓ _parse_critical_lines()      → failed_items + status
  ↓ _detect_equip_errors()       → equip_errors（7种设备/通信错误模式）

file_bk/env_config.yml
  ↓ _parse_env_config()          → instruments（VISA地址 + EQP_ID）

*_MEASUREMENT_Zillnk.json
  ↓ _read_measurement_json()     → first_fail_desc + result（B3B40专有）

Failed_points_*.txt              → 失败项文本摘要（RRU/Apricot专有）

↓ _match_rules()                 → fault_type（优先级：设备错误 > 失败项名 > 日志关键词）
↓ Ollama LLM（增强版可选）       → 提升未分类故障识别率
↓ fault_db.add_record()          → 写入 fault_database.db
```

### 8.3 `CRITICAL - <string>` 日志格式（双产品通用）

```
{timestamp} - CRITICAL - <string> - {TestItemName}, data={value}({unit}), limit=[{lsl},{usl}], result=Pass/Fail
```

示例：
```
2026-04-03 08:29:46 - CRITICAL - <string> - PA CURR CHECK ORBI BandA CH0, data=122.25(mA), limit=[90, 120], result=Fail
2025-10-11 19:08:48 - CRITICAL - <string> - ZBOOT, data=1.1.5, limit=['1.1.4', '1.1.4'], result=Fail
```

### 8.4 故障分类优先级

```
1. 检测到设备/通信错误（equip_errors）→ 优先归为"设备类故障"
2. 失败测试项名称匹配规则             → 归为对应子系统故障
3. 全日志关键词匹配规则               → 兜底分类
4. Ollama LLM分析（增强版）           → 进一步识别未分类故障
```

### 8.5 用户故障知识积累机制

**入口**：菜单 `工具 → 加载故障关系描述文件…`

**推荐工作流**：
1. 运行 `fail_only` 或 `all_with_fail` 分析，得到 `rule_suggestions_*.yaml`
2. 用文本编辑器打开，填写未分类条目的 `fault_type` 和 `suggestion`
3. 菜单导入 → 系统自动合并（相同 keywords 更新，新增追加）
4. 下次分析立即生效，未分类比例持续下降

**YAML 文件格式**（手工编写或基于模板修改）：
```yaml
rules:
  - keywords: "PAM接触,接触异常"
    fault_type: "夹具接触不良"
    suggestion: "检查测试夹具金手指，清洁后重测"
  - keywords: "RX NF ORBI 3600,底噪偏高"
    fault_type: "LNA/滤波器问题"
    suggestion: "对比pass记录NF值，若delta>3dB考虑LNA器件损坏"
```

**导入行为**：
- 相同 keywords → 更新 fault_type 和 suggestion（取最新）
- 不同 keywords → 追加新规则
- 不覆盖全表，支持多次导入累积

### 8.6 fail_only 模式故障模式统计（`_generate_fail_patterns()`）

```python
{
  'total_fail':           int,            # 处理的失败记录总数
  'top_failed_items':     [(item, n)...], # 高频失败测试项 TOP20
  'top_equip_errors':     [(label, n)...],# 高频设备/通信错误 TOP10
  'top_fault_types':      [(type, n)...], # 故障分类分布
  'unclassified_samples': [str...],       # 未分类故障的 first_fail_desc（最多20条）
}
```

`unclassified_samples` 中的描述直接来自 ATE 的 `FirstFailCaseDescription`，是**添加新规则的最直接线索**。

### 8.7 分析结束后自动生成的报告文件

**`generate_fault_barcode_list(db_path, output_path)`**  
从 fault_database.db 查询所有 fail/unknown 记录，写入 `fault_barcodes.xlsx`。  
列：条码、完整条码、工站、机台、测试时间、状态、故障类型、失败测试项数量、首次失败描述。  
按测试时间倒序排列。

**`generate_rule_suggestions_yaml(db_path, output_path)`**  
调用 `_generate_fail_patterns()` 汇总失败数据，生成 `rule_suggestions_<时间戳>.yaml`。  
内容三部分：
1. 未分类故障样本（最优先，无任何规则匹配）
2. 高频失败测试项 Top 10（附出现次数）
3. 设备/通信错误参考（注释行）

工程师填写后通过菜单导入即可更新规则库。

---

## 九、已知产品类型

| 产品 | 目录特征 | 条码格式 | 关键文件 |
|------|---------|---------|---------|
| B3/B40 PA板 | `TestResult/ORBI_B3(40)/FT1/X11_X11/` | 双条码 WVxxxxxx_WVxxxxxx | `*_MEASUREMENT_Zillnk.json`（含FirstFailDesc） |
| Apricot RRU | 拷贝结构，无TestResult层，直接 `FT1/R1B/` | 单条码 BFxxxxxxx | `Failed_points_*.txt`；截图目录以测试项命名 |

---

## 十、文件结构

```
cpk_analyse_platform/
  main.py                        GUI主窗口（LocalAnalysisTab + PlaceholderTab×2 + 菜单）
                                 _get_product_category()、_infer_product_name()
                                 StationRow、MergeRuleRow、LocalAnalysisTab
                                 CPKAnalysisPlatform（主窗口类，F11全屏）
  gen_ppt.py                     架构说明PPT生成脚本（基于公司模板，9页，红白配色）
  core/
    data_extractor.py            run_extraction(mode)          — latest_pass/all/fail_only
                                 run_extraction_all_pass()     — 全部pass记录遍历
                                 run_extraction_traverse()     — folder_direct全量遍历（pass+fail）
                                 generate_folder_direct_excel() — 3-Sheet失败分析Excel
                                 generate_missing_report()     — 缺失条码Excel报表
                                 generate_duplicate_report()   — 重复条码Excel报表
                                 discover_barcodes()           — 无条码列表时自动发现
                                 _walk_all_records_in_folder() — pass+fail全量递归扫描
                                 _walk_all_pass_in_folder()    — 仅pass记录递归扫描
                                 _read_fail_items_from_xlsx()  — 读取xlsx中的失败测试项
                                 check_has_direct_files()      — 检测文件夹是否直接含目标文件
    cpk_calculator.py            analyze_xlsx_folder()、analyze_json_folder()
                                 _file_time_from_name()（文件名14位时间戳优先排序）
                                 values存(barcode, value, is_pass)三元组
    html_report.py               generate_report()，CPK专项报告（自包含，无外部依赖）
    html_comprehensive_report.py generate_comprehensive_report()，6 Tab综合报告（Chart.js）
    html_fail_report.py          generate_fail_report()，失败分析报告（帕累托+汇总卡片）
    fault_db.py                  SQLite持久层（fault_rules/fault_records/fault_stats）
                                 init_db()、add_rule()、update_rule()、get_rules()
                                 add_record()、get_cross_station_barcodes()
    fault_analyzer.py            run_fault_analysis(mode='all'/'fail_only')
                                 generate_fault_barcode_list()（故障条码Excel）
                                 generate_rule_suggestions_yaml()（规则建议YAML模板）
  all_with_fail_design.html      all_with_fail模式设计方案文档
  requirements.txt               Python依赖
  REQUIREMENTS.md                本需求文档
  app_config.json                用户配置持久化（自动生成）
```

---

## 十一、配置持久化（app_config.json）

```json
{
  "excel_path": "...",
  "out_dir": "...",
  "stations": [
    {"type": "FT1", "folder": "E:/..."}
  ],
  "cpk_mode": "最后一次pass数据",
  "fault_level": "基础版（规则库）",
  "file_type": "xlsx",
  "merge_rules": [
    {"src": "FT2", "dst": "FT1"}
  ]
}
```

---

## 十二、待实现功能

### P1（高价值）
- [ ] HTML 报告增强：工站Tab → 测试项Tab两级结构；CPK列默认隐藏（输入`cpk`显示）；产品数据检索Tab（三层展开：条码→测试项→子项明细）
- [ ] `env_comp/*.csv` 校准漂移检测：跨时间比较同路径补偿值，检测链路插损变化
- [ ] `test_limits.yml` 跨批次哈希比对：检测限值变更时间点与失败率突变关联

### P2（中优先级）
- [ ] Pass/Fail 差异比对引擎：同条码跨次测量值 delta 分析（见 `all_with_fail_design.html`）
- [ ] 用户知识输入后端 `load_knowledge_yaml()`（菜单入口已就绪，后端函数待独立封装）
- [ ] 时间窗口聚类：同一工站1小时内 ≥3 个不同DUT同一项失败 → 仪器健康告警
- [ ] 共失败关联性矩阵：发现高关联失败项对
- [ ] `all_with_fail` 跨站推断自动化（目前输出跨站条码列表，需人工判断）
- [ ] 工站合并逻辑扩展至 `all_pass` / `all_with_fail` / `fail_only` 模式（目前仅 `folder_direct` 模式有效）

### P3（后期）
- [ ] 截图内容AI分析（需Ollama视觉模型）
- [ ] 增强版Ollama提示词优化（注入diff结果、历史案例few-shot）
- [ ] `barcode_comparisons` 表（pass/fail比对结果持久化，见设计文档）

### 功能二/三（规划中）
- [ ] 功能二：深科技 MES 导出数据 CPK 分析（待样本数据）
- [ ] 功能三：立讯 MES 导出数据 CPK 分析（待样本数据）
