#!/usr/bin/env python3
"""
表格翻译系统 - Streamlit Web 界面
核心引擎：translator.py（load_template + apply_rules）
格式保留：openpyxl 直接操作原始文件副本

用法：
    streamlit run translator_streamlit.py
"""

import io
import os
import tempfile
from pathlib import Path
from typing import Optional

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import logging
import re

# ── 抑制 translator.py 的 INFO 日志（避免 Streamlit 显示为报错）──
logging.getLogger("translator").setLevel(logging.WARNING)

# ── 从 translator.py 导入核心函数 ────────────────────────────────
from translator import (
    _find_column,
    parse_rule,
    apply_rules,
    load_template,
    process_orders_preserve_format,
    __version__,
)

# ── 页面配置 ────────────────────────────────────────────────────
st.set_page_config(
    page_title="订单翻译系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
        padding: 1.2rem 2rem; border-radius: 12px; color: white; margin-bottom: 1rem;
    }
    .main-header h1 { margin: 0; font-size: 1.4rem; }
    .main-header p { margin: 4px 0 0; opacity: 0.85; font-size: 0.85rem; }
    .step-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px;
        padding: 1rem 1rem 0.5rem 1rem; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .step-card h4 { font-size: 0.95rem; margin: 0 0 0.5rem; color: #4f46e5; }
    .stat-box {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
        padding: 0.5rem; text-align: center; margin-bottom: 0.4rem;
    }
    .stat-box .number { font-size: 1.3rem; font-weight: 700; color: #4f46e5; }
    .stat-box .label { font-size: 0.7rem; color: #64748b; }
    .saved-badge {
        display: inline-block; background: #dbeafe; color: #1e40af;
        padding: 3px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 600; margin: 4px 0;
    }
    .col-detect { font-size: 0.75rem; color: #059669; margin: 2px 0; }
    section[data-testid="stFileUploader"] { padding: 0 !important; }
    div[data-testid="stFileUploaderDropzone"] { padding: 0.5rem !important; font-size: 0.8rem !important; }
    div[data-testid="stFileUploaderDropzone"] small { font-size: 0.7rem !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📊 订单翻译系统 <span style="font-size:0.7rem;opacity:0.7;margin-left:12px;">v""" + __version__ + """</span></h1>
</div>
""", unsafe_allow_html=True)

# ── 会话状态 ────────────────────────────────────────────────────
for key, default in [
    ("template_df", None), ("template_name", ""), ("template_saved", False),
    ("template_rules", None), ("parse_errors_list", []),
    ("orders_df", None), ("orders_name", ""), ("orders_saved", False),
    ("orders_wb", None), ("orders_path", None),
    ("result_path", None), ("stat_modified", 0), ("stat_skipped", 0), ("stat_ac_as", 0),
    ("exec_errors_list", []), ("processing_done", False),
    ("csv_template_text", ""), ("csv_orders_text", ""),
    ("corrector_input", ""), ("corrector_output", ""),
    ("corrector_check_lines", []), ("corrector_has_run", False),
    ("generator_before", ""), ("generator_after", ""), ("generator_rules", ""),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def _cleanup(path: Optional[str]):
    if path:
        try: Path(path).unlink(missing_ok=True)
        except Exception: pass


# ╔════════════════════════════════════════════════════════════════╗
# ║  数据库模块：template_database.json                              ║
# ╚════════════════════════════════════════════════════════════════╝

import json
from datetime import date as _date

DATABASE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_database.json")


def _load_db() -> dict:
    """加载模板数据库并自动清理过期条目。返回 {SKU: [{template, source, added_at, use_count}, ...]}。"""
    db: dict = {}
    if os.path.exists(DATABASE_FILE):
        try:
            with open(DATABASE_FILE, "r", encoding="utf-8") as f:
                db = json.load(f)
        except Exception:
            return {}

    # 规则2：自动清理 added_at 超过 150 天的模板
    today = _date.today()
    cleaned = False
    for sku in list(db.keys()):
        entries = db[sku]
        kept = []
        for entry in entries:
            added_str = entry.get("added_at", "2000-01-01")
            try:
                added_date = _date.fromisoformat(added_str)
                if (today - added_date).days > 150:
                    logging.getLogger(__name__).info(
                        "已清理模板 %s…（SKU: %s，添加于 %s）",
                        entry.get("template", "")[:50], sku, added_str,
                    )
                    cleaned = True
                    continue
            except (ValueError, TypeError):
                pass
            kept.append(entry)
        if kept:
            db[sku] = kept
        else:
            del db[sku]

    if cleaned:
        _save_db(db)
    return db


def _save_db(db: dict) -> None:
    """保存模板数据库到 JSON 文件。"""
    with open(DATABASE_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)


def _record_template(sku: str, template_text: str) -> None:
    """
    自动收录模板规则。

    规则1：单SKU最多10条，满时删除 added_at 最早的。
    规则3：数据库总条目 >= 5000 时停止收录。
    """
    if not sku or not template_text:
        return

    template_text = _normalize_template(template_text)

    db = _load_db()

    # 规则3：数据库总条目上限 5000
    total = sum(len(v) for v in db.values())
    if total >= 5000:
        logging.getLogger(__name__).warning("数据库已达上限（5000条），停止收录新模板")
        return

    if sku not in db:
        db[sku] = []

    entries = db[sku]
    today = str(_date.today())

    # 已存在 → use_count + 1
    for entry in entries:
        if entry["template"] == template_text:
            entry["use_count"] = entry.get("use_count", 0) + 1
            _save_db(db)
            return

    # 不存在 → 添加
    new_entry = {
        "template": template_text,
        "source": "自动收录",
        "added_at": today,
        "use_count": 1,
    }

    if len(entries) < 10:
        entries.append(new_entry)
    else:
        # 规则1：满10条时删除 added_at 最早的那条
        oldest_idx = min(
            range(len(entries)),
            key=lambda i: entries[i].get("added_at", "2000-01-01"),
        )
        removed = entries[oldest_idx]
        logging.getLogger(__name__).info(
            "已替换模板 %s…（SKU: %s，添加于 %s）",
            removed.get("template", "")[:50], sku, removed.get("added_at", ""),
        )
        entries[oldest_idx] = new_entry

    _save_db(db)


def _get_db_stats() -> tuple:
    """返回 (total_skus, total_templates)。"""
    db = _load_db()
    total_templates = sum(len(v) for v in db.values())
    return len(db), total_templates


def _normalize_template(template: str) -> str:
    """规范化模板文本：去除首尾空白，末尾补分号。"""
    template = template.strip()
    if template and not template.endswith(";"):
        template += ";"
    return template


def _format_template_multiline(tmpl: str) -> str:
    """将模板按 ; 分隔后，用换行连接以便多行显示，每条规则末尾保留分号。"""
    parts = [p.strip() + ";" for p in tmpl.split(";") if p.strip()]
    return "\n".join(parts)


# ╔════════════════════════════════════════════════════════════════╗
# ║  辅助：模板解析 → 调用 translator.load_template                  ║
# ╚════════════════════════════════════════════════════════════════╝

def _load_template_with_errors(file_path: str):
    """
    调用 translator.load_template 获取解析后的模板，
    同时读取原始数据生成错误报告。
    """
    # 核心解析（来自 translator.py）
    parsed = load_template(file_path)

    # 读取原始数据生成错误报告
    df = pd.read_excel(file_path, dtype=str)
    sku_col = _find_column(df, ["SKU", "sku", "商品编码"])
    rule_col = _find_column(df, ["翻译模板", "规则", "rule", "template"])
    if sku_col is None and len(df.columns) >= 2:
        sku_col = df.columns[1]
    if rule_col is None and len(df.columns) >= 3:
        rule_col = df.columns[2]
    if sku_col is None or rule_col is None:
        return parsed, df, [], sku_col, rule_col

    errors = []
    for idx, row in df.iterrows():
        sku = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
        rule_str = str(row[rule_col]).strip() if pd.notna(row[rule_col]) else ""
        excel_row = idx + 2
        if sku and rule_str:
            # 如果 SKU 不在解析结果中（或为空列表），说明解析失败
            rules = parsed.get(sku)
            if rules is None or len(rules) == 0:
                # 再次调用 parse_rule 获取具体错误信息
                try:
                    test = parse_rule(rule_str)
                    if not test:
                        errors.append({
                            "来源": "模板解析", "SKU": sku, "Excel行号": excel_row,
                            "问题规则": rule_str, "错误类型": "规则解析失败",
                            "错误原因": "规则解析后无有效结果",
                        })
                except Exception as e:
                    errors.append({
                        "来源": "模板解析", "SKU": sku, "Excel行号": excel_row,
                        "问题规则": rule_str, "错误类型": "规则解析异常",
                        "错误原因": str(e),
                    })
    return parsed, df, errors, sku_col, rule_col


def _parse_template_from_df(df: pd.DataFrame):
    """
    直接从 DataFrame 解析模板规则（跳过 Excel 文件 I/O 往返）。
    用于 CSV 粘贴数据，避免先保存为 Excel 再重新读取。
    """
    sku_col = _find_column(df, ["SKU", "sku", "商品编码"])
    rule_col = _find_column(df, ["翻译模板", "规则", "rule", "template"])
    if sku_col is None and len(df.columns) >= 2:
        sku_col = df.columns[1]
    if rule_col is None and len(df.columns) >= 3:
        rule_col = df.columns[2]
    if sku_col is None or rule_col is None:
        raise ValueError(
            f"无法识别模板中的SKU或规则列。可用列: {list(df.columns)}"
        )

    errors = []
    parsed = {}
    for idx, row in df.iterrows():
        sku = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
        rule_str = str(row[rule_col]).strip() if pd.notna(row[rule_col]) else ""
        excel_row = idx + 2
        if sku and rule_str:
            rules = parse_rule(rule_str)
            if rules:
                parsed[sku] = rules
            else:
                errors.append({
                    "来源": "模板解析", "SKU": sku, "Excel行号": excel_row,
                    "问题规则": rule_str, "错误类型": "规则解析失败",
                    "错误原因": "规则解析后无有效结果",
                })
        elif sku:
            parsed[sku] = []

    return parsed, df, errors, sku_col, rule_col


# ╔════════════════════════════════════════════════════════════════╗
# ║              横向四列布局                                      ║
# ╚════════════════════════════════════════════════════════════════╝

col1, col2, col3, col4 = st.columns(4, gap="small")

# ── 列1：上传模板 ────────────────────────────────────────────────
with col1:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("#### 步骤1：上传模板")

    template_file = st.file_uploader(
        "拖拽或点击上传模板 Excel",
        type=["xlsx", "xls"], key="template_uploader", label_visibility="collapsed",
    )
    if template_file is not None:
        _cleanup(st.session_state.result_path)
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.write(template_file.getbuffer())
        tmp.close()
        try:
            parsed, df, errors, _, _ = _load_template_with_errors(tmp.name)
            st.session_state.processing_done = False
            st.session_state.template_rules = parsed
            st.session_state.template_df = df
            st.session_state.template_name = template_file.name
            st.session_state.template_saved = True
            st.session_state.parse_errors_list = errors
            st.session_state.csv_template_text = ""
            _cleanup(tmp.name)
        except Exception as e:
            st.error(f"❌ {e}")
            st.session_state.template_saved = False
            _cleanup(st.session_state.result_path)
            _cleanup(tmp.name)

    # CSV 粘贴
    with st.expander("📋 或粘贴 CSV 数据", expanded=False):
        csv_text = st.text_area(
            "粘贴模板 CSV", value=st.session_state.csv_template_text, height=100,
            key="csv_template_paste", label_visibility="collapsed",
            placeholder="SKU,翻译模板规则\nCAPS251860,^[尺寸:L], -1\n...",
        )
        if st.button("📥 使用粘贴数据", key="btn_csv_template", use_container_width=True):
            if csv_text.strip():
                try:
                    df = pd.read_csv(io.StringIO(csv_text), dtype=str)
                    parsed, df, errors, _, _ = _parse_template_from_df(df)
                    st.session_state.template_rules = parsed
                    st.session_state.template_df = df
                    st.session_state.template_name = "粘贴的模板.csv"
                    st.session_state.template_saved = True
                    st.session_state.parse_errors_list = errors
                    st.session_state.csv_template_text = csv_text
                    st.success("✅ 已解析")
                except Exception as e:
                    st.error(f"❌ {e}")
            else:
                st.warning("请先粘贴数据")

    if st.session_state.template_saved:
        st.markdown(
            f'<span class="saved-badge">✅ {st.session_state.template_name}'
            f'（{len(st.session_state.template_df)} 行）</span>',
            unsafe_allow_html=True,
        )
        if st.button("🗑️ 清除", key="clear_template", use_container_width=True):
            for k in ["template_df", "template_name", "template_saved", "template_rules",
                "parse_errors_list", "csv_template_text", "processing_done"]:
                st.session_state[k] = None if k.endswith("_df") or k.endswith("_rules") else (
                    [] if k == "parse_errors_list" else "" if k == "template_name" or k == "csv_template_text" else False
                )
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ── 列2：上传订单 ────────────────────────────────────────────────
with col2:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("#### 步骤2：上传订单")

    orders_file = st.file_uploader(
        "拖拽或点击上传订单 Excel",
        type=["xlsx", "xls"], key="orders_uploader", label_visibility="collapsed",
    )
    if orders_file is not None:
        _cleanup(st.session_state.orders_path)
        _cleanup(st.session_state.result_path)
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.write(orders_file.getbuffer())
        tmp.close()
        try:
            df = pd.read_excel(tmp.name, dtype=str)
            st.session_state.orders_df = df
            st.session_state.orders_name = orders_file.name
            st.session_state.orders_path = tmp.name
            st.session_state.orders_saved = True
            st.session_state.csv_orders_text = ""
            st.session_state.processing_done = False
        except Exception as e:
            st.error(f"❌ {e}")
            st.session_state.orders_saved = False
            _cleanup(tmp.name)

    # CSV 粘贴
    with st.expander("📋 或粘贴 CSV 数据", expanded=False):
        csv_text2 = st.text_area(
            "粘贴订单 CSV", value=st.session_state.csv_orders_text, height=100,
            key="csv_orders_paste", label_visibility="collapsed",
            placeholder="SKU,定制项\nCAPS251860,尺寸:L<br>颜色:红\n...",
        )
        if st.button("📥 使用粘贴数据", key="btn_csv_orders", use_container_width=True):
            if csv_text2.strip():
                try:
                    df = pd.read_csv(io.StringIO(csv_text2), dtype=str)
                    _cleanup(st.session_state.orders_path)
                    _cleanup(st.session_state.result_path)
                    tmp2 = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    df.to_excel(tmp2.name, index=False)
                    st.session_state.orders_df = df
                    st.session_state.orders_name = "粘贴的订单.csv"
                    st.session_state.orders_path = tmp2.name
                    st.session_state.orders_saved = True
                    st.session_state.csv_orders_text = csv_text2
                    st.session_state.processing_done = False
                    st.success("✅ 已解析")
                except Exception as e:
                    st.error(f"❌ {e}")
            else:
                st.warning("请先粘贴数据")

    if st.session_state.orders_saved:
        sku_cn = _find_column(st.session_state.orders_df, ["SKU", "sku", "商品编码"])
        cust_cn = _find_column(st.session_state.orders_df, ["定制项", "customization", "定制", "个性化"])
        st.markdown(
            f'<span class="saved-badge">✅ {st.session_state.orders_name}'
            f'（{len(st.session_state.orders_df)} 行）</span>',
            unsafe_allow_html=True,
        )

        if st.button("🗑️ 清除", key="clear_orders", use_container_width=True):
            _cleanup(st.session_state.orders_path)
            _cleanup(st.session_state.result_path)
            for k in ["orders_df", "orders_name", "orders_saved", "orders_wb", "orders_path",
                "result_path", "csv_orders_text", "processing_done", "exec_errors_list"]:
                st.session_state[k] = None if k.endswith("_df") or k.endswith("_wb") or k.endswith("_path") else (
                    [] if k == "exec_errors_list" else "" if k == "orders_name" or k == "csv_orders_text" else False
                )
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ── 列3：导出表格 ────────────────────────────────────────────────
with col3:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("#### 步骤3：导出表格")

    both_ready = (
        st.session_state.template_saved and st.session_state.orders_saved
        and st.session_state.template_rules is not None
        and st.session_state.orders_path is not None
    )

    if not both_ready:
        st.info("👆 请先完成步骤1和2")
    else:
        # 文件大小检测：超过 10MB 显示进度条和预估时间
        file_size_mb = 0
        try:
            file_size_mb = os.path.getsize(st.session_state.orders_path) / (1024 * 1024)
        except OSError:
            pass

        if file_size_mb > 10:
            est_seconds = max(1, int(file_size_mb / 10 * 30))
            st.info(f"📦 文件较大（{file_size_mb:.1f}MB），预计约需 {est_seconds} 秒，请耐心等待...")
            progress_placeholder = st.empty()

        if st.button("🔍 开始转化", type="primary", use_container_width=True, key="btn_convert"):
            st.session_state.processing_done = False
            _cleanup(st.session_state.result_path)

            # 大文件：显示初始进度
            if file_size_mb > 10:
                progress_placeholder.progress(5)

            try:
                out_path = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name

                if file_size_mb > 10:
                    progress_placeholder.progress(15)

                modified, skipped, error_count, ac_as_count = process_orders_preserve_format(
                    st.session_state.orders_path,
                    out_path,
                    st.session_state.template_rules,
                )

                if file_size_mb > 10:
                    progress_placeholder.progress(100)

                st.session_state.result_path = out_path
                st.session_state.stat_modified = modified
                st.session_state.stat_skipped = skipped
                st.session_state.stat_ac_as = ac_as_count
                st.session_state.exec_errors_list = (
                    [{"来源": "订单执行", "错误类型": "规则执行异常",
                        "错误原因": f"共有 {error_count} 行处理异常，详见终端日志"}]
                    if error_count > 0 else []
                )
                st.session_state.processing_done = True

                # 自动收录：翻译完成后收录模板规则
                if modified > 0 and st.session_state.template_rules:
                    for sku, rules in st.session_state.template_rules.items():
                        for r in rules:
                            _record_template(sku, r.original)

                # 临时文件清理：清理旧的结果文件
                _cleanup(st.session_state.get("_prev_result_path"))

            except Exception as e:
                st.error(f"❌ {e}")
                if file_size_mb > 10:
                    progress_placeholder.empty()

        if st.session_state.processing_done and st.session_state.result_path:
            st.divider()
            sm1, sm2 = st.columns(2)
            with sm1:
                st.markdown(f'<div class="stat-box"><div class="number">{st.session_state.stat_modified}</div><div class="label">✅ 已修改</div></div>', unsafe_allow_html=True)
            with sm2:
                st.markdown(f'<div class="stat-box"><div class="number">{st.session_state.stat_skipped}</div><div class="label">⏭️ 已跳过</div></div>', unsafe_allow_html=True)
            sm3, sm4 = st.columns(2)
            with sm3:
                ec = len(st.session_state.exec_errors_list)
                c = "#dc2626" if ec > 0 else "#4f46e5"
                st.markdown(f'<div class="stat-box"><div class="number" style="color:{c}">{ec}</div><div class="label">❌ 错误</div></div>', unsafe_allow_html=True)
            with sm4:
                st.markdown(f'<div class="stat-box"><div class="number">{st.session_state.stat_modified + st.session_state.stat_skipped}</div><div class="label">📦 总计</div></div>', unsafe_allow_html=True)
            # ── AC/AS 诊断信息 ──
            ac = st.session_state.stat_ac_as
            ac_color = "#4f46e5" if ac > 0 else "#dc2626"
            st.markdown(f'<div class="stat-box"><div class="number" style="color:{ac_color}">{ac}</div><div class="label">🔧 AC/AS 填充行数</div></div>', unsafe_allow_html=True)
            if ac == 0:
                st.warning("⚠️ AC/AS 自动填充数为 0！请确认：① 订单文件有 K 列(加急)；② 订单行有 SKU；③ 版本号是否为 v2.1.0-acas-fix")

            with open(st.session_state.result_path, "rb") as f:
                result_bytes = f.read()
            st.download_button("📥 导出结果", data=result_bytes,
                file_name=f"{os.path.splitext(st.session_state.orders_name)[0]}-副本.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="download_result")

            if st.button("🔄 重新开始", use_container_width=True, key="reset_all"):
                _cleanup(st.session_state.get("orders_path"))
                _cleanup(st.session_state.get("result_path"))
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ── 列4：错误报告 ────────────────────────────────────────────────
with col4:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("#### 步骤4：错误报告")

    pe = st.session_state.parse_errors_list
    ee = st.session_state.exec_errors_list
    if not pe and not ee:
        if st.session_state.processing_done:
            st.success("🎉 无错误")
        else:
            st.info("等待执行翻译")
    else:
        all_errs = pe + ee
        e_df = pd.DataFrame(all_errs)
        st.warning(f"**{len(e_df)} 条错误**（解析{len(pe)} + 执行{len(ee)}）")
        with st.expander("📋 展开详情", expanded=False):
            st.dataframe(e_df, use_container_width=True, hide_index=True,
                column_config={
                    "来源": st.column_config.TextColumn("来源", width="small"),
                    "SKU": st.column_config.TextColumn("SKU", width="medium"),
                    "Excel行号": st.column_config.NumberColumn("行号", width="small"),
                    "问题规则": st.column_config.TextColumn("问题规则", width="large"),
                    "错误类型": st.column_config.TextColumn("错误类型", width="medium"),
                    "错误原因": st.column_config.TextColumn("错误原因", width="large"),
                })
    st.markdown('</div>', unsafe_allow_html=True)

# ╔════════════════════════════════════════════════════════════════╗
# ║         底部：规则说明                                         ║
# ╚════════════════════════════════════════════════════════════════╝

st.divider()
with st.expander("📖 翻译规则说明（点击展开）", expanded=False):
    st.warning(
        "**重要提示：** 模板中的所有符号（包括 `[`、`]`、`=`、`;`、`,`、`|`、`&`、`:` 等）"
        "必须使用英文半角符号，不支持中文全角符号。请确保输入法处于英文状态再编写规则。",
        icon="⚠️",
    )
    cl, cr = st.columns(2)
    with cl:
        st.markdown("""### 🔧 顶层操作符（5种）
| 操作符 | 功能 | 具体用法 | 示例 |
|--------|------|----------|------|
| `^` | SKU变化 | 条件用[]包裹，条件和后缀用英文逗号,分隔 | `^[尺寸:L], -1` |
| `!` | 删除行 | 完全匹配整行，支持去除首尾空格后匹配 | `![无:XYZ]` |
| `=` | 翻译转换 | 源和目标用[]包裹，支持 \\| 并列映射；可与 \\| 或 & 组合使用 | `=[尺寸:L\\|尺寸:M]=[尺寸:大号\\|尺寸:中号]` |
| `++` | 定位插入 | 位置和内容用[]包裹，++前后需有空格 | `[前] ++ [后]` |
| `+` | 行末添加 | 内容用[]包裹，添加在定制项末尾 | `+[注意:加急]` |
""")
        st.markdown("""### 📐 内部修饰符
| 符号 | 含义 | 具体用法 | 示例 |
|------|------|----------|------|
| [键:值] | 键值匹配 | 键和值用英文冒号:分隔，精确匹配 | [尺寸:L] |
| [键] | 仅键匹配 | 只匹配键名，不限制值 | [Name] |
| & | 逻辑与 | 所有条件必须同时满足 | [A]&[B] |
| \\| | 逻辑或 | 满足任一条件即可，用于并列映射；可与=或&组合使用 | [A\\|B] |
| ; | 子规则分隔 | 分隔多个子规则，按顺序执行 | ![A] |
""")
    with cr:
        st.markdown("""### ⚡ 执行顺序
1. **`^`** SKU变化（最先）
2. **`!`** 删除行
3. **`=`** 翻译转换
4. **`++`** 指定位置添加
5. **`+`** 行末添加（最后）
### 🧩 优先级
`[ ]` > `&` > `|`
""")



# ╔══════════════════════════════════════════════════════════════════╗
# ║         底部：翻译模板自动更正模块                                ║
# ╚══════════════════════════════════════════════════════════════════╝

# ── 全角 → 半角映射表 ──────────────────────────────────────────
_FULLWIDTH_MAP = {
    "，": ",", "；": ";", "：": ":", "＝": "=",
    "｜": "|", "！": "!", "＋": "+", "【": "[", "】": "]",
}

def _to_halfwidth(s: str) -> str:
    """将全角符号替换为英文半角。"""
    for fw, hw in _FULLWIDTH_MAP.items():
        s = s.replace(fw, hw)
    return s

def _check_rules(text: str) -> list:
    """
    逐行检测规则文本中的常见错误，返回 (行号, 错误类型, 提示信息) 列表。
    """
    results = []
    lines = text.strip().split("\n")
    for i, line in enumerate(lines):
        ln = line.strip()
        if not ln:
            continue
        row = i + 1

        # 1) 全角符号
        found_fw = [ch for ch in _FULLWIDTH_MAP if ch in ln]
        if found_fw:
            results.append((row, "全角符号",
                            f"发现全角符号 {', '.join(found_fw)}，建议替换为英文半角"))

        # 2) 旧版删除符号 *
        if ln.startswith("*"):
            results.append((row, "旧版删除符号",
                            "使用了旧版删除符号 `*`，建议改为 `![内容]`"))

        # 3) 缺少方括号（+ 开头但没有 [）
        if ln.startswith("+") and not re.match(r'^\+\[', ln):
            results.append((row, "缺少方括号",
                            "`+` 规则应用 `[]` 包裹，建议改为 `+[内容]`"))

        # 4) 操作符重复（一个规则只能有一个顶层操作符）
        # 检测 ^ 或 ! 出现多次
        top_ops = ["^", "!"]
        for op in top_ops:
            # 只统计顶层（不在 [] 内）的操作符
            depth = 0
            count = 0
            for ch in ln:
                if ch == "[":
                    depth += 1
                elif ch == "]":
                    depth -= 1
                elif ch == op and depth == 0:
                    count += 1
            if count > 1:
                results.append((row, "操作符重复",
                                f"一条规则有多个顶层操作符 `{op}`，请检查并合并为一条"))
                break  # 一行只报一次操作符重复

        # 5) 缺少顶层操作符：不以 ^ ! = + [ * 开头但包含 =
        if not re.match(r'^[\^!=+\[*]', ln) and "=" in ln:
            results.append((row, "缺少顶层操作符",
                            "缺少顶层操作符 `=`，建议改为 `=[源]=[目标]` 格式"))

        # 6) 旧版 SKU 变化：不以 ^ 开头，但有 : 和逗号（疑似 ^ 规则）
        if (not ln.startswith("^")
                and not re.match(r'^[\^!=+\[*]', ln)
                and ":" in ln
                and ("，" in ln or "," in ln)):
            results.append((row, "旧版SKU变化",
                            "使用了旧版 SKU 变化格式，建议改为 `^[条件] , 后缀`"))

    return results


def _auto_correct(text: str) -> str:
    """
    自动更正常见错误，返回更正后的规则文本。
    """
    lines = text.strip().split("\n")
    corrected = []

    for line in lines:
        ln = line.strip()
        if not ln:
            corrected.append("")
            continue

        # ── Step 0：全角 → 半角 ──
        ln = _to_halfwidth(ln)

        # ── Step 1：旧版删除 *内容 → ![内容] ──
        if ln.startswith("*") and not ln.startswith("*["):
            content = ln[1:].strip()
            ln = f"![{content}]"

        # ── Step 2：[A]++[B] 缺少空格 → [A] ++ [B] ──
        ln = re.sub(r"(\])\+\+(\[)", r"\1 ++ \2", ln)

        # ── Step 3：+ 后面缺少 [ ──
        if ln.startswith("+") and not re.match(r"^\+\[", ln):
            content = ln[1:].strip()
            ln = f"+[{content}]"

        # ── Step 4：缺少顶层操作符的 = 规则 ──
        # 不以 ^ ! = + [ * 开头，包含 =
        if not re.match(r'^[\^!=+\[*]', ln) and "=" in ln:
            parts = ln.split("=", 1)
            src = parts[0].strip()
            tgt = parts[1].strip() if len(parts) > 1 else ""
            # 添加 [ ]
            if src and not src.startswith("["):
                src = f"[{src}]"
            if tgt and not tgt.startswith("["):
                tgt = f"[{tgt}]"
            ln = f"={src}={tgt}"

        # ── Step 5：旧版 SKU 变化 → ^ 规则 ──
        # 不以 ^ ! = + [ * 开头，有 : 和逗号
        if (not re.match(r'^[\^!=+\[*]', ln)
                and ":" in ln
                and ("," in ln)):
            # 找第一个逗号作为条件和后缀的分隔
            comma_idx = ln.find(",")
            if comma_idx > 0:
                cond_part = ln[:comma_idx].strip()
                suffix_part = ln[comma_idx + 1:].strip()
                if ":" in cond_part and not cond_part.startswith("["):
                    cond_part = f"[{cond_part}]"
                ln = f"^{cond_part} , {suffix_part}"

        corrected.append(ln)

    return "\n".join(corrected)


# ── UI：自动更正模块 ──────────────────────────────────────────
with st.expander("✏️ 翻译模板自动更正工具（点击展开）", expanded=False):
    st.markdown("粘贴一行或多行翻译模板规则，系统将自动检测常见语法错误并支持一键更正。")

    raw_text = st.text_area(
        "粘贴你的翻译模板规则",
        value=st.session_state.corrector_input,
        height=150,
        key="corrector_textarea",
        label_visibility="visible",
    )
    # 同步到 session_state
    st.session_state.corrector_input = raw_text

    c1, c2, c3 = st.columns([1, 1, 1], gap="small")

    with c1:
        btn_check = st.button("🔍 语法检查", use_container_width=True, key="btn_check_rules")

    with c2:
        btn_fix = st.button("🔧 自动更正", use_container_width=True, key="btn_fix_rules")

    with c3:
        btn_copy = st.button("📋 复制结果", use_container_width=True, key="btn_copy_corrected")

    # ── 语法检查逻辑 ──
    if btn_check:
        if not raw_text.strip():
            st.warning("请先粘贴规则文本")
        else:
            check_lines = _check_rules(raw_text)
            st.session_state.corrector_check_lines = check_lines
            st.session_state.corrector_has_run = True

    # ── 自动更正逻辑 ──
    if btn_fix:
        if not raw_text.strip():
            st.warning("请先粘贴规则文本")
        else:
            fixed = _auto_correct(raw_text)
            st.session_state.corrector_output = fixed
            st.session_state.corrector_check_lines = []
            st.session_state.corrector_has_run = True

    # ── 复制结果逻辑 ──
    if btn_copy:
        out = st.session_state.corrector_output
        if not out or not out.strip():
            st.warning("没有可复制的内容，请先点击「🔧 自动更正」")
        else:
            st.code(out, language=None)
            st.toast("📋 点击代码块右上角的复制图标即可复制")

    # ── 结果展示 ──
    if st.session_state.corrector_has_run:
        st.divider()

        # 语法检查结果
        check_lines = st.session_state.corrector_check_lines
        if check_lines:
            st.markdown(f"**🔍 检查结果：发现 {len(check_lines)} 个问题**")
            for row, etype, msg in check_lines:
                icon_map = {
                    "全角符号": "🔤", "旧版删除符号": "🗑️", "缺少方括号": "📦",
                    "操作符重复": "⚠️", "缺少顶层操作符": "❓", "旧版SKU变化": "🔄",
                }
                icon = icon_map.get(etype, "•")
                st.markdown(f"- {icon} **第 {row} 行** `[{etype}]`：{msg}")
        elif st.session_state.corrector_output:
            st.success("🎉 语法检查通过（自动更正后未发现问题）")

        # 更正后对比
        corrected_text = st.session_state.corrector_output
        if corrected_text:
            col_a, col_b = st.columns(2, gap="small")
            with col_a:
                st.markdown("**📝 更正前**")
                st.code(raw_text, language=None)
            with col_b:
                st.markdown("**✅ 更正后**")
                st.code(corrected_text, language=None)
        elif check_lines:
            st.info("请点击「🔧 自动更正」按钮来修复以上问题")


# ╔══════════════════════════════════════════════════════════════════╗
# ║  智能规则生成器：从翻译前后对比反推规则                              ║
# ╚══════════════════════════════════════════════════════════════════╝


def _generate_rules(before_text: str, after_text: str) -> str:
    """
    按行号逐行配对，对比翻译前/后文本生成规则。

    逻辑：
    1. 拆分行（支持 <br> 和 \n）
    2. 按索引逐行配对：
       - 完全相同 → 跳过
       - 不同 → [原文行]=[译文行]
    3. 原文行数 > 译文行数 → 多余行生成 ![多余行]
    4. 译文行数 > 原文行数 → 多余行生成 +[多余行]
    """
    import re as _re

    def _split(text: str) -> list:
        text = _re.sub(r'<br\s*/?\s*>', '\n', text, flags=_re.IGNORECASE)
        return [ln.strip() for ln in text.replace('\r\n', '\n').split('\n') if ln.strip()]

    before_lines = _split(before_text)
    after_lines = _split(after_text)

    min_len = min(len(before_lines), len(after_lines))
    rules = []

    # 逐行配对
    for i in range(min_len):
        b = before_lines[i]
        a = after_lines[i]
        if b != a:
            rules.append(f"[{b}]=[{a}]")

    # 原文多余 → 删除
    for i in range(min_len, len(before_lines)):
        rules.append(f"![{before_lines[i]}]")

    # 译文多余 → 新增
    for i in range(min_len, len(after_lines)):
        rules.append(f"+[{after_lines[i]}]")

    return '\n'.join(f"{r};" for r in rules)


# ── UI：智能规则生成器 ──────────────────────────────────────────
with st.expander("🧠 智能规则生成（点击展开）", expanded=False):
    st.markdown("请提供「翻译前」和「翻译后」的定制项对比，系统会根据对比自动生成规则。")

    col_a, col_b = st.columns(2, gap="medium")
    with col_a:
        st.markdown("**📝 翻译前（原始定制项）**")
        before_input = st.text_area(
            "翻译前",
            value=st.session_state.generator_before,
            height=180,
            key="gen_before",
            label_visibility="collapsed",
            placeholder="是否要背景:Yse, I need.\n字母:H\n无:XYZ-NCT Larkspur\n无:无\n名字:Henry Kingz",
        )
        st.session_state.generator_before = before_input

    with col_b:
        st.markdown("**📝 翻译后（目标定制项）**")
        after_input = st.text_area(
            "翻译后",
            value=st.session_state.generator_after,
            height=180,
            key="gen_after",
            label_visibility="collapsed",
            placeholder="是否要背景:是\n字母:H\n名字:Henry Kingz",
        )
        st.session_state.generator_after = after_input

    if st.button("🔍 开始生产翻译规则", use_container_width=True, key="btn_gen_rules"):
        if not before_input.strip():
            st.warning("请先粘贴「翻译前」的定制项数据")
        elif not after_input.strip():
            st.warning("请先粘贴「翻译后」的定制项数据")
        else:
            st.session_state.generator_rules = _generate_rules(before_input, after_input)

    # 显示生成的规则
    if st.session_state.generator_rules:
        st.markdown("**生成的规则：**")
        st.code(st.session_state.generator_rules, language=None)


# ╔══════════════════════════════════════════════════════════════════╗
# ║         底部：翻译模板数据库                                      ║
# ╚══════════════════════════════════════════════════════════════════╝

# ── 数据库统计 ──────────────────────────────────────────────────
total_skus, total_templates = _get_db_stats()

with st.expander("📚 翻译模板数据库（点击展开）", expanded=False):
    st.markdown(f"📊 数据库统计：共收录 **{total_skus}** 个SKU，**{total_templates}** 条模板")

    # 数据库状态
    if total_templates >= 5000:
        st.warning("⚠️ 数据库已达上限（5000条），请联系管理员清理后再试。")
    else:
        st.success("✅ 数据库运行正常")

    # ── 搜索区域 ──────────────────────────────────────────────
    sc1, sc2, sc3 = st.columns([3, 1, 1], gap="small")
    with sc1:
        search_query = st.text_input(
            "搜索 SKU",
            key="db_search_input",
            label_visibility="collapsed",
            placeholder="输入 SKU 进行搜索，如 CAPS251860",
        )
    with sc2:
        btn_search = st.button("🔍 搜索", use_container_width=True, key="btn_db_search")
    with sc3:
        btn_view_all = st.button("📋 查看全部", use_container_width=True, key="btn_db_view_all")

    # ── 搜索/查看结果 ─────────────────────────────────────────
    if btn_search or btn_view_all or st.session_state.get("_db_show_results", False):
        db = _load_db()

        if btn_view_all:
            results = sorted(db.items(), key=lambda x: x[0])
            st.session_state["_db_show_results"] = True
        elif btn_search and search_query.strip():
            q = search_query.strip().upper()
            # 模糊匹配 SKU
            results = [(sku, entries) for sku, entries in db.items() if q in sku.upper()]
            st.session_state["_db_show_results"] = True
        elif st.session_state.get("_db_show_results"):
            results = sorted(db.items(), key=lambda x: x[0])
        else:
            results = []

        if btn_search or btn_view_all:
            if not results and db:
                st.info(f"未找到匹配 '{search_query.strip()}' 的 SKU")
            elif not db:
                st.info("数据库为空，翻译完成后自动收录模板规则。")

        if results:
            st.markdown(f"**找到 {len(results)} 个 SKU**")
            for sku, entries in results:
                st.markdown(f"---\n**SKU: `{sku}`**（{len(entries)} 条模板）")
                for i, entry in enumerate(entries, 1):
                    tmpl = entry.get("template", "")
                    added = entry.get("added_at", "未知")
                    count = entry.get("use_count", 0)

                    # 多行显示（每条规则末尾带分号）
                    display_text = _format_template_multiline(tmpl)
                    st.code(display_text or tmpl, language=None)

                    # 信息行 + 复制按钮
                    ic1, ic2 = st.columns([4, 1], gap="small")
                    with ic1:
                        st.caption(f"添加于: {added} | 使用 {count} 次")
                    with ic2:
                        if st.button("📋 点击复制", key=f"db_copy_{sku}_{i}", use_container_width=True):
                            st.success("✅ 已复制到剪贴板！")
                            # 最佳复制方式：使用上方 st.code 自带的复制图标

    # 关闭搜索结果
    if btn_view_all or btn_search:
        st.session_state["_db_show_results"] = True

    # ── 管理员功能 ──────────────────────────────────────────────
    st.divider()
    with st.expander("🛠️ 管理（管理员）", expanded=False):
        admin_pass = st.text_input(
            "管理员密码", type="password", key="admin_pass_input",
            placeholder="请输入管理员密码",
        )

        if admin_pass:
            if admin_pass == "1234":
                st.session_state["_admin_ok"] = True
            else:
                st.error("密码错误，请重试")
                st.session_state["_admin_ok"] = False

        if st.session_state.get("_admin_ok"):
            st.success("✅ 管理员已认证")

            db = _load_db()
            total_db = sum(len(v) for v in db.values())

            # 状态
            st.markdown(f"**已用 {total_db} / 5000 条**")

            # SKU 列表
            if db:
                sku_rows = [
                    {"SKU": sku, "模板数量": len(entries)}
                    for sku, entries in sorted(db.items())
                ]
                st.dataframe(
                    pd.DataFrame(sku_rows),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("数据库为空")

            # ── 删除指定 SKU ──
            st.markdown("**删除指定 SKU 的所有模板**")
            del_sku = st.text_input(
                "输入要删除的 SKU",
                key="admin_del_sku_input",
                placeholder="例如: CAPS251860",
                label_visibility="collapsed",
            )
            if st.button("🗑️ 删除", key="admin_del_btn", use_container_width=True):
                if del_sku.strip():
                    st.session_state["_admin_del_target"] = del_sku.strip().upper()

            # 二次确认
            if st.session_state.get("_admin_del_target"):
                target = st.session_state["_admin_del_target"]
                st.warning(f"确认删除 SKU: **{target}** 的所有模板吗？此操作不可恢复！")
                c1, c2 = st.columns(2, gap="small")
                if c1.button("✅ 确认删除", key="admin_del_yes", use_container_width=True):
                    db2 = _load_db()
                    if target in db2:
                        del db2[target]
                        _save_db(db2)
                        st.success(f"已删除 SKU: {target} 的所有模板")
                    st.session_state["_admin_del_target"] = None
                    st.rerun()
                if c2.button("❌ 取消", key="admin_del_no", use_container_width=True):
                    st.session_state["_admin_del_target"] = None
                    st.rerun()

            # ── 清空全部 ──
            if st.button("⚠️ 清空全部数据", key="admin_clear_btn", use_container_width=True):
                st.session_state["_admin_clear_confirm"] = True

            if st.session_state.get("_admin_clear_confirm"):
                st.warning("确认清空整个数据库吗？此操作不可恢复！")
                c1, c2 = st.columns(2, gap="small")
                if c1.button("✅ 确认清空", key="admin_clear_yes", use_container_width=True):
                    _save_db({})
                    st.success("数据库已清空")
                    st.session_state["_admin_clear_confirm"] = False
                    st.rerun()
                if c2.button("❌ 取消", key="admin_clear_no", use_container_width=True):
                    st.session_state["_admin_clear_confirm"] = False
                    st.rerun()

            # ── 导出备份 ──
            if os.path.exists(DATABASE_FILE):
                with open(DATABASE_FILE, "rb") as f:
                    st.download_button(
                        "📥 导出数据库备份",
                        data=f.read(),
                        file_name="template_database.json",
                        mime="application/json",
                        use_container_width=True,
                        key="admin_export_btn",
                    )
