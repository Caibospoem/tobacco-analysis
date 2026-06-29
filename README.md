# 卷烟货源投放策略自动分析

每周上传新的 xlsx 货源投放策略表，GitHub Actions 自动解析并输出 JSON。

## 目录结构

```
.
├── .github/workflows/
│   └── analyze-cigarettes.yml      # GitHub Actions 工作流
├── scripts/
│   └── cigarette_parser.py          # 解析脚本（可复用 CLI 模块）
├── data/                            # ← 每周上传 xlsx 到这里
│   └── 全区卷烟货源投放策略表（...）.xlsx
├── output/                          # ← 自动生成的 JSON 存这里
│   └── 卷烟投放策略_2026年6月21日-26日.json
├── requirements.txt                 # Python 依赖
└── README.md
```

## 工作流程

```
你上传新 xlsx 到 data/ 并 git push
        │
        ▼
GitHub Actions 检测到 data/**/*.xlsx 变更
        │
        ▼
检出代码 → 安装 Python + openpyxl
        │
        ▼
对比 git diff 找出本次变更的 xlsx 文件
        │
        ▼
逐个调用 cigarette_parser.py 解析
        │
        ▼
生成的 JSON 提交回 output/ 目录
        │
        ▼
同时上传为 Artifact（保留 90 天，便于下载）
```

## 部署步骤

### 1. 创建 GitHub 仓库

新建一个仓库（公开或私有均可），把本模板所有文件推上去：

```bash
git init
git add .
git commit -m "初始化卷烟投放策略自动分析项目"
git branch -M main
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git push -u origin main
```

### 2. 配置仓库权限（关键）

工作流需要提交文件回仓库，必须开启写权限：

> 仓库 **Settings** → **Actions** → **General** → 滚动到 **Workflow permissions**
> → 选择 **Read and write permissions** → 保存

### 3. 上传 xlsx 文件

每周把新的 xlsx 文件放到 `data/` 目录并推送：

```bash
# 方式一：命令行
cp "全区卷烟货源投放策略表（2026年6月28日下午-2026年7月3日上午）.xlsx" data/
git add data/
git commit -m "data: 上传第N期投放策略表"
git push

# 方式二：网页上传
# 在 GitHub 仓库页面进入 data/ 目录 → Add file → Upload files
```

### 4. 查看结果

- **自动触发**：push 后约 1-2 分钟，工作流开始运行
- **查看 JSON**：运行完成后，`output/` 目录会自动出现新的 JSON 文件
- **下载产物**：在工作流运行详情页的 **Artifacts** 区域可下载 JSON 压缩包
- **查看运行日志**：仓库 **Actions** 标签页 → 点击对应运行记录

## 手动触发（调试用）

如果只想分析某个文件而不通过 push，可手动触发：

> 仓库 **Actions** 标签页 → 左侧选 **分析卷烟投放策略** → 右侧 **Run workflow**
> → 在 `file` 输入框填入文件名（如 `xxx.xlsx`，留空则分析 data/ 下所有 xlsx）→ Run

## JSON 输出结构

```json
{
  "投放时间": "2026年6月21日-26日",
  "单位": "条",
  "数据来源": "xxx.xlsx",
  "档级投放": {
    "档位汇总": {
      "30档": { "卷烟数量": 121, "卷烟列表": [...] },
      ...
      "1档": { ... }
    },
    "条件投放策略": [ ... ]
  },
  "标签投放": [
    { "策略编号": "1", "策略说明": "...", "卷烟数量": 4, "卷烟列表": [...] },
    ...
  ],
  "雪茄投放": {
    "档位投放": { "A档": [...], "B档": [...], ... },
    "标签投放": [ ... ]
  }
}
```

## 本地调试

```bash
pip install -r requirements.txt
python scripts/cigarette_parser.py --input data/xxx.xlsx --output-dir output/
```

## 故障排查

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| 工作流不触发 | xlsx 不在 `data/` 目录 | 确保文件放在 `data/` 下 |
| JSON 未提交回仓库 | 仓库未开启写权限 | Settings → Actions → Workflow permissions → Read and write |
| 解析报错 | xlsx 格式与预期不符 | 查看运行日志，确认三个工作表名（档级投放/标签投放/雪茄投放）存在 |
| 中文文件名乱码 | git 未配置 UTF-8 | 运行 `git config --global core.quotepath false` |
| 首次提交 diff 失败 | 无历史记录 | 脚本已兼容：diff 失败时自动分析 data/ 下所有 xlsx |

## 技术说明

- **触发机制**：`on.push.paths` 过滤 `data/**/*.xlsx`，仅当 xlsx 变更时触发，避免无关提交浪费运行
- **变更检测**：`git diff HEAD~1 HEAD` 对比本次推送变更的文件，只分析新文件，效率高
- **幂等性**：重复分析同一文件会覆盖同名 JSON，不会产生重复
- **日期提取**：从 xlsx 单元格 A2 自动提取投放时间，作为 JSON 文件名的一部分，便于按期归档
