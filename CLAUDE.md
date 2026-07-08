# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

A Chinese mortgage calculator (房贷计算器) web app built with Streamlit. Supports first/second home purchases with three loan types: pure provident fund (纯公积金), pure commercial (纯商贷), and combined (组合贷).

## Commands

```bash
# Install
pip install -r requirements.txt

# Run locally
streamlit run app.py

# Run on a custom port
streamlit run app.py --server.port 8502
```

## Architecture

Single-file app: `app.py`. No database, no external APIs — all calculations happen client-side in Python.

### Interest rate rules (hardcoded)

| Scenario | 公积金 | 商贷 |
|---|---|---|
| 首套房 | 2.6% | 3.0% |
| 二套房 | 2.6% | 3.3% |

### Down payment minimums

- 公积金参与（纯公积金 / 组合贷）→ ≥20%
- 纯商贷 → ≥15%

### Calculation engine

- Uses **等额本息** (equal monthly installment): `M = P × r × (1+r)^n / ((1+r)^n - 1)`
- Combined loans: provident fund portion allocated first (up to configurable cap), remainder goes to commercial loan
- All currency values stored in 元 internally; UI displays 万元 for large amounts and 元 for monthly payments

## Deployment

Recommended: **Streamlit Community Cloud** (free).

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo, select `app.py` as the entry point
4. Deploy — no extra config needed
