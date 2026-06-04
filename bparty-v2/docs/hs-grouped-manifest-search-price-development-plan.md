# HS 归并与互联网搜价清关流程开发计划

日期：2026-06-04

## 背景

任务 `CMDUCHN3357475` 暴露出当前清关生成流程的核心问题：系统没有先整理客户清单，而是主要依赖提单品名和替换表补齐输出，导致客户反馈品名挑选、编码、数量、重量、单价、货值都不够可信。

最新共识：

- 客户清单先按 HS 编码作为唯一主键做归并整理。
- 后续选品优先来自归并后的客户清单，不足时再从替换表补。
- 重量不仅要总重量闭合，还要单件重量、单箱重量、每箱数量合理。
- 单价只做普通互联网搜索，不接入 Amazon Associates、eBay API、Walmart API、BestBuy API、Keepa、SerpAPI 或任何需要 API key 的第三方服务。
- 互联网搜索得到的零售价格样本换算成单件零售价，再按 30% 得到申报参考单价。
- 计划后续将整个软件迁移到海外服务器，便于访问搜索引擎、电商页面和 codeflagai；用户已用海外 VPN 验证 codeflagai 访问速度可接受。

## 重要执行约束

### 服务器代码为准

正式开发前必须先确认服务器代码和本地代码是否同步。若不一致，以服务器 `47.114.75.48:/opt/bparty-v2` 的代码为准。

开发开始前的前置步骤：

1. 在本地确认当前分支、工作区和远端状态。
2. SSH 到服务器 `root@47.114.75.48`，检查 `/opt/bparty-v2` 是否为 Git 工作区。
3. 若服务器是 Git 工作区：
   - 比较服务器 `git rev-parse HEAD` 与本地 `git rev-parse HEAD`。
   - 比较服务器 `git status --short` 与本地工作区。
   - 如服务器有未提交改动，先拉取或打包服务器版本到本地，作为本地开发基线。
4. 若服务器不是 Git 工作区：
   - 用 `rsync --dry-run` 对比服务器 `/opt/bparty-v2` 与本地 `bparty-v2`。
   - 下载服务器快照到本地临时目录。
   - 以服务器快照覆盖或合并本地 `bparty-v2` 后，再创建开发分支。
5. 只有确认基线后，才创建本地分支，例如 `codex/hs-grouped-search-price-flow`。

### 不接第三方付费/API Key 服务

价格模块只允许使用：

- 普通网页搜索结果。
- 可公开访问的电商页面。
- 客户清单原始单价。
- 本地历史价格缓存。
- LLM 估算作为最后兜底。

明确不做：

- Amazon Associates / PA-API / Creators API。
- eBay Browse API。
- Walmart API。
- BestBuy API。
- Keepa。
- SerpAPI、DataForSEO、Bright Data、Oxylabs 等需要账号或 API key 的服务。

## 目标流程

新流程：

```text
原始清单 Excel
  -> 原始行解析
  -> 按 HS 归并 canonical_manifest
  -> 为每个 HS group 选择代表品名
  -> 清单 HS group 优先选品
  -> 不足时替换表补充
  -> codeflagai 校验 HS/税率/认证
  -> 普通互联网搜索估价
  -> 重量/数量/价格硬校验
  -> 税金目标优化
  -> 输出清关 Excel + 审计摘要
```

## 数据结构设计

### `ManifestHsGroup`

新增结构，用 HS 作为唯一键归并客户清单：

```python
@dataclass
class ManifestHsGroup:
    hs: str
    source_rows: list[int]
    zh_names: list[str]
    en_names: list[str]
    materials: list[str]
    usages: list[str]
    total_ctns: float
    total_qty: float
    total_real_weight: float
    total_gross_weight: float
    total_declared_value: float
    unit_price_values: list[float]
    canonical_zh: str
    canonical_en: str
    representative_material: str
    representative_usage: str
    warnings: list[str]
```

归并规则：

- HS 标准化到 10 位；无法标准化的行进入异常列表，不直接参与候选。
- 同 HS 的行合并数量、箱数、重量、货值。
- 同 HS 下保留所有原始品名，不丢 source_rows。
- 同 HS 下品名跨度过大时不拆 key，但标记 `warnings`，后续代表品名选择要更谨慎。
- 代表品名优先选择权重更高、数量更多、语义更常见、英文更可搜索的名称。

### `PriceEvidence`

新增互联网搜价证据结构：

```python
@dataclass
class PriceEvidence:
    query: str
    samples: list[dict]
    retail_unit_price: float
    declared_unit_price: float
    basis: str
    confidence: float
    source: str
```

样本字段：

```json
{
  "title": "Plastic hair clips set 100 pcs",
  "url": "https://...",
  "price_usd": 6.99,
  "pack_qty": 100,
  "unit_price_usd": 0.0699,
  "source_domain": "example.com"
}
```

## 模块改造计划

### 1. 清单归并模块

位置建议：优先放在 `engine.py` 内，稳定后可拆到 `manifest_normalizer.py`。

新增函数：

- `build_manifest_hs_groups(manifest: ManifestSummary) -> list[ManifestHsGroup]`
- `normalize_manifest_hs(value) -> str`
- `choose_group_representative_name(group: ManifestHsGroup) -> tuple[str, str]`
- `build_manifest_group_candidates(groups) -> list[ProductCandidate]`

改造点：

- 移除当前 `qualified_manifest = []` 和 `crawler_manifest_products skipped` 的跳过逻辑。
- `ProductCandidate.source` 增加或使用 `manifest_group`。
- `ProductCandidate.source_rows` 写入该 HS group 的所有原始行号。
- `ProductCandidate.source_label` 标记为 `客户清单/HS归并`。

验收：

- `CMDUCHN3357475` 清单的 185 行应归并成明显少于原始行数的 HS groups。
- 同 HS 重复行的数量、箱数、重量、货值应被合计。
- 输出候选池中应出现来自清单的发饰、钥匙扣、塑料饰品等相关 HS group。

### 2. 清单优先选品

新增函数：

- `select_candidates_manifest_first(manifest_candidates, bill, replacement_pool, options)`
- `map_bill_products_to_hs_groups(bill, manifest_groups)`
- `summarize_replacement_usage(selected, target_item_count)`

规则：

- 先从清单 HS group 中选。
- 提单品名必须覆盖，但优先映射到清单已有 HS group。
- 替换表只补不足行数，原清单里能用多少就用多少。
- 替换表比例只作为审计指标记录，不作为硬失败条件。
- 替换品不能与清单主语义明显无关。

验收：

- `CMDUCHN3357475` 输出大部分行来自清单归并池。
- 替换表只补清单候选不足的行数，并记录替换比例。
- 不再为了补行数引入明显无关品类。

### 3. HS 与 codeflagai 校验

复用现有 `StrictTaxCrawler`，但查询对象从原始 185 行变成 HS group 候选。

规则：

- 优先查 HS。
- HS 查不到或税率/认证不可用时，再用代表品名 + 材质查。
- 清单 HS 与 codeflagai 匹配品名明显冲突时，进入候选过滤，不直接输出。
- 低税率不能覆盖语义不匹配。

新增或加强：

- `validate_hs_semantic_match(candidate, tax_data)`
- `candidate.tax_match_source` 必须写清楚来自 HS 查询还是品名查询。
- `filter_summary` 记录 HS 冲突、认证冲突、语义不匹配。

验收：

- 编码错误、多个编码粘连、爬虫品名明显不一致时能阻断。
- 输出每行都能解释采用 HS 的理由。

### 4. 普通互联网搜价模块

新增文件建议：`bparty-v2/price_search.py`

职责：

- 构造普通搜索 query。
- 抓取公开网页或搜索结果页面。
- 提取标题、价格、包装数量。
- 换算单件零售价。
- 过滤异常样本。
- 计算申报参考价。
- 缓存结果。

搜索 query 策略：

- `{canonical_en} price`
- `{canonical_en} retail price`
- `{material} {canonical_en} price`
- `site:walmart.com {canonical_en} price`
- `site:target.com {canonical_en} price`
- `site:ebay.com {canonical_en} price`

价格解析规则：

- 识别 `$6.99`、`USD 6.99` 等美元价格。
- 识别 `100 pcs`、`100 count`、`pack of 12`、`12 pack`、`set of 50`。
- 无法识别包装数量时默认 `pack_qty=1`，但降低 confidence。
- 过滤二手、批发整箱、明显不相关标题。
- 使用中位数或截尾均值，不用简单平均。
- `declared_unit_price = retail_unit_price * 0.3`。

兜底顺序：

1. 普通互联网搜索价格。
2. 历史价格缓存。
3. 客户清单原始单价的归并中位数。
4. LLM 估算。

验收：

- 不需要任何 API key。
- 价格证据写入任务记录。
- 单价不能为了贴近税金被压到无依据低价。

### 5. 重量与数量硬校验

现有代码已经有合理范围和 warning，但要把不合理单件/单箱重量从 warning 升级为 hard failure。

规则：

- 总重量必须闭合。
- 总箱数必须闭合。
- 数量必须大于等于箱数。
- 数量必须是箱数整数倍。
- 单件重量必须在合理范围。
- 单箱重量必须在合理范围。
- 每箱数量必须在合理范围。
- 若总重量闭合会导致某行不合理，不能输出，只能重新选候选或重算数量/重量。

新增或加强：

- `validate_weight_plausibility_hard(rows, selected)`
- `validate_quantity_consistency_with_manifest_group(rows, selected)`

验收：

- 不再出现“连衣裙单件 1000g”这类结果。
- 发夹、钥匙扣等小件的单件重量必须落在轻小件范围。

### 6. 价格与货值硬校验

新增：

- `attach_price_evidence_to_candidates(selected)`
- `validate_price_plausibility(rows, selected)`
- `validate_value_floor(rows, manifest)`

规则：

- 单价必须落在价格证据合理区间附近。
- 若只有低置信度估算，输出审计摘要必须标明。
- 总货值不能低到原始货值的极端低比例。
- 税金优化只能在价格合理区间内做，不能突破价格证据。

建议初始参数：

- 申报单价默认为互联网零售单件价的 30%。
- 可允许浮动区间为 20%-45%，用于贴近税金。
- 总货值低于原清单货值 10% 时默认阻断，除非后续业务确认要放宽。

验收：

- `CMDUCHN3357475` 不再出现原始货值 182,859.56 USD、输出货值只有 5,882.30 USD 这种约 3.22% 的极低结果。

### 7. 审计摘要

任务记录中新增 `audit` 字段，前端可后续展示。

每行输出记录：

- 来源：客户清单 HS group / 替换表 / 提单。
- source_rows。
- 原始 HS。
- 采用 HS。
- codeflagai 匹配结果。
- 税率和认证。
- 重量依据。
- 价格证据。
- 替换原因。
- 风险提示。

验收：

- 客户问“为什么选这个品名/编码/重量/单价”时，可以直接从任务记录解释。

## 测试计划

### 单元测试

新增或扩展 `bparty-v2/tests/test_engine.py`：

- HS 归并：同 HS 多行合并。
- HS 异常：空 HS、粘连 HS、长度异常。
- 代表品名选择：同 HS 多品名时保留 source_names 并选择可搜索名称。
- 替换补足：原清单候选能用多少就用多少，替换比例只记录不阻断。
- 重量硬校验：单件/单箱重量超范围阻断。
- 价格解析：从标题和文本解析价格、pack qty、unit price。
- 价格兜底：搜索无样本时 fallback 到清单价或 LLM 估算。

### 回归测试

以 `CMDUCHN3357475` 为固定回归样例：

- 输出行大部分来自客户清单 HS group。
- 替换表只补清单候选不足的行数，并记录替换比例。
- 发夹和钥匙扣数量不能无依据放大。
- 单件重量和单箱重量全部通过硬校验。
- 总货值不能低到原始货值的 3%。
- 每行有 price evidence 或明确 fallback basis。

### 真实链路测试

在海外网络环境或海外服务器上验证：

- codeflagai 登录与查询稳定。
- 普通互联网搜索可访问。
- 价格搜索超时不会拖垮任务。
- 缓存命中时不重复搜价。

## 分阶段执行计划

### 阶段 0：同步基线

- 检查服务器 `/opt/bparty-v2` 与本地 `bparty-v2`。
- 以服务器代码为准同步本地。
- 确认本地工作区干净。
- 创建本地分支 `codex/hs-grouped-search-price-flow`。

### 阶段 1：HS 归并与清单候选池

- 实现 `ManifestHsGroup`。
- 实现 HS group 构建。
- 用 HS group 替代当前跳过的 manifest candidate 流程。
- 增加单元测试。

### 阶段 2：清单优先选品与替换补足审计

- 实现 manifest-first selection。
- 实现 replacement usage audit。
- 让 `CMDUCHN3357475` 至少大部分输出来自清单。

### 阶段 3：重量与数量硬约束

- 将单件重量、单箱重量、每箱数量改为硬校验。
- 重量闭合失败时阻断，不只写 warning。
- 增加回归测试。

### 阶段 4：普通互联网搜价

- 新增 `price_search.py`。
- 实现普通搜索、页面解析、价格样本清洗、缓存、兜底。
- 将 `PriceEvidence` 写入 candidate 和 task record。

### 阶段 5：价格/货值硬约束与税金优化

- 单价只能在价格证据允许区间内调整。
- 总货值不作为硬约束，仅进入审计摘要。
- 税金目标在合理范围内尽量贴近，最终税金不得超过目标税金上浮 10%，不能突破重量和价格常理。

### 阶段 6：审计摘要与前端兼容

- 在任务 JSON 中保留审计摘要。
- 先不强制改前端展示，但保证 API payload 可用。
- 如需要，再做前端审计详情展示。

## 风险与待确认

- HS 作为唯一主键会把同 HS 下多个相近品类归到一起，代表品名选择必须可解释。
- 普通互联网搜索没有 API，页面结构和反爬不稳定，需要超时、缓存和 fallback。
- 海外服务器访问 codeflagai 已由用户用 VPN 初步验证，但正式迁移前仍需在目标服务器 smoke test。
- 总货值下限 10% 已取消；当前硬规则只控制最终税金不得超过目标税金上浮 10%。
- 价格申报比例 30% 已确认作为默认逻辑，但浮动范围 20%-45% 是否接受仍需确认。

## 用户确认后立即执行的第一批动作

1. 连接 `47.114.75.48`，检查 `/opt/bparty-v2` 与本地代码差异。
2. 若服务器代码更新，先同步服务器代码到本地。
3. 创建本地分支 `codex/hs-grouped-search-price-flow`。
4. 先做阶段 1 和阶段 2，并跑单元测试。
5. 阶段 1/2 可运行后，再继续阶段 3-5。
