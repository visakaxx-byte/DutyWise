# BParty V2

独立的乙方清关资料生成服务。输入清单 Excel、运输提单 PDF、期望税金和最终条目数，经过 codeflagai 智能归类/税率查询、规则筛选和税金分配后，输出 `Commercial Invoice&Packing List`。

## 内容

- `app.py`：FastAPI 服务和静态前端入口
- `engine.py`：清单/提单解析、候选筛选、税金分配、Excel 输出
- `crawler_client.py`：codeflagai 登录和智能归类接口
- `llm_client.py`：OpenAI-compatible chat completions 调用
- `static/`：上传处理页面
- `templates/清关模板.xlsx`：输出模板
- `docs/海关编码查找.xlsx`：常用替换清单，优先读取 `常用1`，再读取 `20260330`
- `rules/`：可维护的允许/禁用品名和认证规则
- `requirements.txt`：Python 依赖
- `.env.example`：运行配置模板
- `deploy/`：systemd 部署文件

## 配置

复制 `.env.example` 为 `.env`，填写真实配置：

```bash
cp .env.example .env
```

必需项：

```text
CRAWLER_USERNAME
CRAWLER_PASSWORD
```

`LLM_*` / `DOUBAO_*` 仍可保留在 `.env` 中，当前税金目标流程不依赖 LLM。

## 本地运行

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

打开：

```text
http://127.0.0.1:8002/
```

如果端口被占用：

```bash
python3 -m uvicorn app:app --host 0.0.0.0 --port 8012
```

## 命令行运行

```bash
python3 run_example.py \
  --manifest "/path/to/manifest.xlsx" \
  --bill "/path/to/bill.pdf" \
  --target-tax-amount 700 \
  --target-item-count 10
```

## 真实链路验证

会真实调用 LLM 和 codeflagai：

```bash
python3 validate_live_flow.py \
  --manifest "/path/to/manifest.xlsx" \
  --bill "/path/to/bill.pdf" \
  --target-tax-amount 700 \
  --target-item-count 10
```

## 离线测试

```bash
make test
make validate-package
```

## 处理规则

- codeflagai 查询优先使用品名 + 材质，客户原始 HTS 只做兜底。
- 基础税率必须 `<20%`；最终预计税金只使用基础税率，不叠加加征税率。
- 认证提示为空可用；只出现 `Lacey Act`、`TSCA` 或二者组合时可用；其他认证/监管提示默认排除。
- 优先使用客户清单里的合格品名，不足目标条目数时从 `docs/海关编码查找.xlsx` 补齐。
- 最终输出行数等于用户输入的目标输出行数，总毛重对齐清单重量总和。
- 输出行会按常用编码表和客户清单派生的单箱重量、单件重量、单价范围做常理约束；无可行解时阻断并返回原因。
- 预计税金在重量和价格常理范围内尽量贴近期望税金，不能为命中税金强行突破常理范围。

## 部署

见 `deploy/README_DEPLOY.md`。
