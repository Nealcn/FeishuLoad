"""Pure, Tk-free helpers for the optional XTF desktop GUI.

The GUI never reimplements synchronization.  It builds a strict YAML v2 file
and runs the existing ``XTF.py sync`` CLI in a subprocess.
"""

from __future__ import annotations

import json
import math
import os
import re
import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Mapping, Sequence
from urllib.parse import parse_qs, urlparse

import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parents[1]

BASE_URL_RE = re.compile(r"/base/([A-Za-z0-9_-]+)")
SHEET_URL_RE = re.compile(r"/sheets/([A-Za-z0-9_-]+)")
TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]{8,}$")

KIND_LABELS = {
    "text": "文本",
    "number": "数字",
    "select": "单选/多选",
    "datetime": "日期",
    "checkbox": "复选框",
    "user": "人员",
    "group_chat": "群组",
    "link": "关联",
    "location": "地理位置",
    "attachment": "附件",
    "formula": "公式",
    "lookup": "查找引用",
    "created_at": "创建时间",
    "updated_at": "修改时间",
    "created_by": "创建人",
    "updated_by": "修改人",
    "auto_number": "自动编号",
    "unsupported": "不支持",
}

MODE_LABELS = {
    "append": "只追加",
    "incremental": "只新增缺失 key",
    "full": "更新 + 新增",
    "overwrite": "删除匹配后重建",
    "clone": "清空后重写",
}

ACTION_LABELS = {
    "create_fields": "创建字段",
    "create_records": "新增记录",
    "update_records": "更新记录",
    "delete_records": "删除记录",
    "clear_range": "清空范围",
    "write_range": "写入范围",
    "append_rows": "追加行",
    "write_columns": "写入列",
    "apply_sheet_config": "字段格式设置",
}

UNIT_LABELS = {
    "field": "字段",
    "record": "记录",
    "row": "行",
    "column": "列",
    "range": "范围",
}

STATUS_LABELS = {
    "success": "成功",
    "noop": "无需写入",
    "planned": "仅预览",
    "failed": "失败",
    "partial": "部分完成",
    "indeterminate": "结果未知",
}

EXIT_LABELS = {
    0: "成功",
    1: "内部错误",
    2: "命令用法错误",
    3: "配置或输入错误",
    4: "认证或权限错误",
    5: "远端读取或计划失败",
    6: "写入失败或部分完成",
    7: "写后确认失败或超时",
    8: "远端结果未知",
    130: "用户中断",
}


class GuiError(RuntimeError):
    """A user-facing error with an optional actionable hint."""

    def __init__(self, message: str, *, hint: str = "") -> None:
        super().__init__(message)
        self.hint = hint


@dataclass(frozen=True)
class TargetRef:
    kind: str
    app_token: str = ""
    table_id: str = ""
    view_id: str | None = None
    spreadsheet_token: str = ""
    sheet_id: str | None = None
    original_url: str = ""

    @property
    def label(self) -> str:
        if self.kind == "bitable":
            return f"Base {self.app_token} / 表 {self.table_id or '未选择'}"
        return f"Sheet {self.spreadsheet_token} / 工作表 {self.sheet_id or '未选择'}"


@dataclass(frozen=True)
class TargetField:
    name: str
    kind: str
    writable: bool
    multiple: bool = False
    raw_type: Any = None
    schema: Any = field(default=None, repr=False, compare=False)

    @property
    def computed(self) -> bool:
        return not self.writable

    @property
    def kind_label(self) -> str:
        if self.kind == "select" and self.multiple:
            return "多选"
        return KIND_LABELS.get(self.kind, self.kind)


@dataclass
class TargetInfo:
    ref: TargetRef
    table_name: str
    fields: list[TargetField]
    records: list[dict[str, Any]]
    backend: str

    @property
    def field_map(self) -> dict[str, TargetField]:
        return {item.name: item for item in self.fields}

    @property
    def record_count(self) -> int:
        return len(self.records)


@dataclass(frozen=True)
class ColumnStat:
    nonempty: int
    unique: int

    @property
    def duplicate(self) -> int:
        return max(0, self.nonempty - self.unique)


@dataclass
class SourceInfo:
    path: Path
    sheet_name: str | None
    frame: pd.DataFrame
    columns: list[str]
    rows: int
    preview: list[list[str]]
    stats: dict[str, ColumnStat]


@dataclass(frozen=True)
class IndexCandidate:
    name: str
    safe: bool
    source_nonempty: int
    source_unique: int
    target_nonempty: int
    target_unique: int
    reason: str


@dataclass(frozen=True)
class LoadedConfig:
    app_id: str
    secret: str | None
    source_path: str
    sheet_name: str | None
    target: TargetRef | None
    mode: str
    index_column: str | None
    selective_columns: list[str] | None
    verify_remote_writes: bool
    create_missing_fields: bool
    field_strategy: str
    backend: str
    batch_size: int


# ---------------------------------------------------------------------------
# URL / source helpers
# ---------------------------------------------------------------------------


def parse_target_url(text: str) -> TargetRef:
    raw = (text or "").strip()
    if not raw:
        raise GuiError("请填写飞书 Base 链接。")
    parsed = urlparse(raw if "://" in raw else "https://" + raw)
    query = parse_qs(parsed.query)
    base_match = BASE_URL_RE.search(parsed.path)
    if base_match:
        return TargetRef(
            kind="bitable",
            app_token=base_match.group(1),
            table_id=(query.get("table") or [""])[0],
            view_id=(query.get("view") or [None])[0],
            original_url=raw,
        )
    sheet_match = SHEET_URL_RE.search(parsed.path)
    if sheet_match:
        return TargetRef(
            kind="sheet",
            spreadsheet_token=sheet_match.group(1),
            sheet_id=(query.get("sheet") or [None])[0],
            original_url=raw,
        )
    if TOKEN_RE.fullmatch(raw):
        return TargetRef(kind="bitable", app_token=raw, original_url=raw)
    raise GuiError(
        "无法识别链接。",
        hint="请粘贴类似 https://xxx.feishu.cn/base/<app_token>?table=<table_id> 的 Base 链接。",
    )


def list_sheet_names(path: str | Path) -> list[str]:
    target = Path(path)
    if target.suffix.lower() not in {".xlsx", ".xls"}:
        return []
    try:
        return list(pd.ExcelFile(target).sheet_names)
    except Exception:
        return []


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    if isinstance(value, Mapping):
        return not value
    if isinstance(value, (list, tuple, set)):
        return not value or all(_is_blank(item) for item in value)
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _display_cell(value: Any) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.isoformat()
    text = str(value)
    return text if len(text) <= 80 else text[:77] + "..."


def read_source(path: str | Path, sheet_name: str | None = None) -> SourceInfo:
    target = Path(path).expanduser()
    if not target.is_file():
        raise GuiError(f"找不到数据文件：{target}")
    from core.reader import DataFileReader

    if not DataFileReader.is_supported(target):
        raise GuiError(
            f"不支持的文件格式：{target.suffix}",
            hint="支持 .xlsx、.xls、.csv（CSV 为实验性支持）。",
        )
    kwargs: dict[str, Any] = {}
    if sheet_name not in (None, "") and target.suffix.lower() in {".xlsx", ".xls"}:
        kwargs["sheet_name"] = sheet_name
    try:
        frame = DataFileReader().read_file(target, **kwargs)
    except Exception as exc:
        raise GuiError(f"读取文件失败：{exc}") from exc

    columns = [str(column).strip() for column in frame.columns]
    if not columns:
        raise GuiError("文件没有表头或没有任何列。")
    if any(not column for column in columns):
        raise GuiError("文件表头中存在空列名。")
    if len(set(columns)) != len(columns):
        raise GuiError("文件表头存在重复列名，无法确定写入字段。")
    frame.columns = columns

    preview = [
        [_display_cell(value) for value in row]
        for row in frame.head(8).itertuples(index=False, name=None)
    ]
    stats: dict[str, ColumnStat] = {}
    for column in columns:
        values = [value for value in frame[column].tolist() if not _is_blank(value)]
        stats[column] = ColumnStat(
            nonempty=len(values),
            unique=len({_display_cell(value) for value in values}),
        )
    return SourceInfo(
        path=target,
        sheet_name=sheet_name or None,
        frame=frame,
        columns=columns,
        rows=len(frame),
        preview=preview,
        stats=stats,
    )


def sanitize_source(
    source: SourceInfo,
    keep_columns: Sequence[str],
    output: str | Path,
) -> Path:
    columns = [column for column in keep_columns if column in source.frame.columns]
    if not columns:
        raise GuiError("没有可写入的列。")
    destination = Path(output)
    destination.parent.mkdir(parents=True, exist_ok=True)
    source.frame.loc[:, columns].to_excel(destination, index=False, engine="openpyxl")
    return destination


# ---------------------------------------------------------------------------
# Feishu probing (read-only)
# ---------------------------------------------------------------------------


def _make_auth(app_id: str, app_secret: str, delay: float = 0.05):
    from api.auth import FeishuAuth
    from api.base import RateLimiter, RetryableAPIClient

    transport = RetryableAPIClient(max_retries=1, rate_limiter=RateLimiter(delay))
    return FeishuAuth(app_id, app_secret, api_client=transport), transport


def _friendly_api_error(exc: Exception) -> GuiError:
    from api.sdk import FeishuAPIError

    if isinstance(exc, FeishuAPIError):
        code = getattr(exc, "code", None)
        hints = {
            99991672: "应用缺少所需权限。请在开发者后台开通对应 scope，并发布新的应用版本。",
            1254302: "应用没有该 Base/表的角色权限。请把应用添加为文档应用，或检查高级权限设置。",
            800030104: "表不存在，或应用没有该表的访问权限。",
            99991661: "认证失败，请检查 App ID / App Secret。",
            99991663: "认证失败，请检查 App ID / App Secret。",
            99991664: "认证失败，请检查 App ID / App Secret。",
            99991668: "认证失败，请检查 App ID / App Secret。",
        }
        return GuiError(
            f"飞书接口错误（{code}）：{exc}",
            hint=hints.get(code, "请检查应用权限、文档协作者和网络连接。"),
        )
    return GuiError(str(exc), hint="请检查网络、权限和凭据。")


def _parse_table_items(items: Any) -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    if not isinstance(items, list):
        return result
    for item in items:
        if not isinstance(item, Mapping):
            continue
        table_id = item.get("id") or item.get("table_id")
        if not isinstance(table_id, str) or not table_id:
            continue
        name = item.get("name") or table_id
        result.append({"table_id": table_id, "name": str(name)})
    return result


def _fetch_v1_tables(auth: Any, url: str) -> list[dict[str, str]]:
    import requests

    from api.sdk import FeishuResponseParser

    result: list[dict[str, str]] = []
    page_token: str | None = None
    for _ in range(100):
        params: dict[str, Any] = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token
        response = requests.get(
            url, headers=auth.get_auth_headers(), params=params, timeout=30
        )
        data = FeishuResponseParser.parse(response)
        payload = data.get("data") or {}
        result.extend(_parse_table_items(payload.get("items")))
        if not payload.get("has_more"):
            break
        next_token = payload.get("page_token")
        if (
            not isinstance(next_token, str)
            or not next_token
            or next_token == page_token
        ):
            break
        page_token = next_token
    return result


def _fetch_v3_tables(auth: Any, url: str) -> list[dict[str, str]]:
    import requests

    from api.sdk import FeishuResponseParser

    result: list[dict[str, str]] = []
    offset = 0
    for _ in range(100):
        response = requests.get(
            url,
            headers=auth.get_auth_headers(),
            params={"limit": 100, "offset": offset},
            timeout=30,
        )
        data = FeishuResponseParser.parse(response)
        payload = data.get("data") or {}
        items = payload.get("tables")
        page = _parse_table_items(items)
        result.extend(page)
        total = payload.get("total")
        offset += len(page)
        if not page or (isinstance(total, int) and offset >= total):
            break
    return result


def _list_tables_with_auth(
    auth: Any, app_token: str, backend: str
) -> list[dict[str, str]]:
    from api.url import encode_path_segment

    encoded = encode_path_segment(app_token)
    v3_url = f"https://open.feishu.cn/open-apis/base/v3/bases/{encoded}/tables"
    v1_url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{encoded}/tables"
    fetchers = (
        [("v3", _fetch_v3_tables, v3_url), ("v1", _fetch_v1_tables, v1_url)]
        if backend == "base_v3"
        else [("v1", _fetch_v1_tables, v1_url), ("v3", _fetch_v3_tables, v3_url)]
    )
    merged: dict[str, dict[str, str]] = {}
    last_error: Exception | None = None
    for _name, fetcher, url in fetchers:
        try:
            for item in fetcher(auth, url):
                merged.setdefault(item["table_id"], item)
        except Exception as exc:  # noqa: BLE001 - try the other API family
            last_error = exc
    if merged:
        return list(merged.values())
    if last_error is not None:
        raise _friendly_api_error(last_error) from last_error
    return []


def list_tables(
    app_id: str,
    app_secret: str,
    app_token: str,
    backend: str = "base_v3",
) -> list[dict[str, str]]:
    auth, _ = _make_auth(app_id, app_secret)
    try:
        return _list_tables_with_auth(auth, app_token, backend)
    except GuiError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise _friendly_api_error(exc) from exc


def probe_bitable(
    app_id: str,
    app_secret: str,
    ref: TargetRef,
    backend: str = "base_v3",
) -> TargetInfo:
    if ref.kind != "bitable":
        raise GuiError(
            "当前界面只支持多维表格（Bitable）。",
            hint="电子表格请暂时使用命令行：python XTF.py sync --config config.yaml。",
        )
    if not ref.table_id:
        raise GuiError("链接中缺少 table 参数，无法确定目标表。")
    auth, transport = _make_auth(app_id, app_secret)
    try:
        from api.bitable_v1 import BitableV1Backend
        from api.bitable_v3 import BaseV3Backend

        backend_obj = (
            BaseV3Backend(auth, transport)
            if backend == "base_v3"
            else BitableV1Backend(auth, transport)
        )
        schemas = backend_obj.list_fields(ref.app_token, ref.table_id)
        result = backend_obj.list_records(ref.app_token, ref.table_id)
    except Exception as exc:  # noqa: BLE001
        raise _friendly_api_error(exc) from exc

    if not result.complete:
        raise GuiError("目标表读取不完整，已停止。")
    if result.ignored_fields or result.record_not_found:
        raise GuiError("目标表读取存在被忽略字段或缺失记录，已停止。")

    table_name = ref.table_id
    try:
        for item in _list_tables_with_auth(auth, ref.app_token, backend):
            if item["table_id"] == ref.table_id:
                table_name = item["name"]
                break
    except Exception:  # noqa: BLE001 - table name is cosmetic
        pass

    fields = [
        TargetField(
            name=schema.name,
            kind=schema.kind.value,
            writable=schema.writable,
            multiple=schema.multiple,
            raw_type=schema.raw_type,
            schema=schema,
        )
        for schema in schemas
    ]
    records = [dict(record.fields) for record in result.records]
    return TargetInfo(
        ref=ref,
        table_name=table_name,
        fields=fields,
        records=records,
        backend=backend,
    )


def index_candidates(source: SourceInfo, target: TargetInfo) -> list[IndexCandidate]:
    """Return source columns that are safe single-column keys for by_key sync."""
    from core.config import TargetType
    from core.converter import DataConverter
    from core.key_policy import KeyPolicy

    converter = DataConverter(TargetType.BITABLE)
    policy = KeyPolicy()
    writable = {item.name: item for item in target.fields if item.writable}
    candidates: list[IndexCandidate] = []
    source_total = len(source.frame)
    target_total = len(target.records)
    if source_total == 0:
        return candidates

    for column in source.columns:
        target_field = writable.get(column)
        if target_field is None or target_field.schema is None:
            continue
        type_code = converter._field_schema_type_code(target_field.schema)
        source_values: list[str] = []
        target_values: list[str] = []
        bad = False
        for value in source.frame[column].tolist():
            if converter._is_empty_value(value):
                continue
            try:
                key = policy.normalize(value, type_code)
            except Exception:  # noqa: BLE001
                bad = True
                break
            if key is not None:
                source_values.append(key.value)
        if not bad:
            for record in target.records:
                value = record.get(column)
                if converter._is_empty_value(value):
                    continue
                try:
                    key = policy.normalize(value, type_code)
                except Exception:  # noqa: BLE001
                    bad = True
                    break
                if key is not None:
                    target_values.append(key.value)

        source_unique = len(set(source_values))
        target_unique = len(set(target_values))
        if bad:
            reason = "存在无法安全归一化的值"
        elif len(source_values) != source_total:
            reason = "源列存在空值"
        elif len(set(source_values)) != len(source_values):
            reason = "源列存在重复值"
        elif len(target_values) != target_total:
            reason = "目标列存在空值"
        elif len(set(target_values)) != len(target_values):
            reason = "目标列存在重复值"
        else:
            reason = "源/目标均唯一且非空"
        safe = reason == "源/目标均唯一且非空"
        candidates.append(
            IndexCandidate(
                name=column,
                safe=safe,
                source_nonempty=len(source_values),
                source_unique=source_unique,
                target_nonempty=len(target_values),
                target_unique=target_unique,
                reason=reason,
            )
        )
    candidates.sort(key=lambda item: (not item.safe, -item.source_unique, item.name))
    return candidates


# ---------------------------------------------------------------------------
# Config building
# ---------------------------------------------------------------------------


def _validate_config(config: Mapping[str, Any]) -> None:
    from xtf_cli.config import validate_v2_document

    try:
        validate_v2_document(config)
    except Exception as exc:  # noqa: BLE001
        raise GuiError(f"生成的配置无效：{exc}") from exc


def build_config(
    *,
    app_id: str,
    source_path: str | Path,
    target: TargetRef,
    mode: str,
    index_column: str | None = None,
    selective_columns: Sequence[str] | None = None,
    verify_remote_writes: bool = True,
    create_missing_fields: bool = True,
    field_strategy: str = "base",
    backend: str = "base_v3",
    batch_size: int = 500,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    if not app_id.strip():
        raise GuiError("请填写 App ID。")
    if target.kind != "bitable":
        raise GuiError("当前界面只支持多维表格（Bitable）。")
    if not target.app_token or not target.table_id:
        raise GuiError("目标 Base 或数据表不完整。")
    if mode not in {"append", "incremental", "full", "overwrite", "clone"}:
        raise GuiError(f"未知同步模式：{mode}")
    if mode in {"incremental", "full", "overwrite"} and not index_column:
        raise GuiError("该模式需要选择一个唯一索引列。")
    if mode in {"append", "clone"} and selective_columns is not None:
        raise GuiError("该模式不能使用选择性列同步。")
    if field_strategy not in {"raw", "base", "auto", "intelligence"}:
        raise GuiError(f"未知字段类型策略：{field_strategy}")
    if backend not in {"base_v3", "bitable_v1"}:
        raise GuiError(f"未知 Bitable 后端：{backend}")

    source = Path(source_path)
    if not source.is_file():
        raise GuiError(f"找不到数据文件：{source}")

    selective = {
        "enabled": False,
        "columns": [],
        "auto_include_index": True,
        "optimize_ranges": True,
        "max_gap_for_merge": 2,
        "preserve_column_order": True,
    }
    sync: dict[str, Any] = {
        "verify_remote_writes": bool(verify_remote_writes),
        "verify_timeout_seconds": 10.0,
        "verify_interval_seconds": 0.5,
    }
    if mode == "append":
        sync.update({"mode": "incremental", "match_strategy": "append_only"})
    elif mode == "clone":
        sync["mode"] = "clone"
    else:
        sync.update(
            {
                "mode": "incremental" if mode == "incremental" else mode,
                "match_strategy": "by_key",
                "index": {
                    "column": index_column,
                    "datetime_granularity": "exact",
                    "timezone": None,
                },
            }
        )
        if selective_columns is not None:
            columns = list(dict.fromkeys([*selective_columns, str(index_column)]))
            if not columns:
                raise GuiError("选择性同步必须至少选择一列。")
            selective.update({"enabled": True, "columns": columns})
    sync["selective"] = selective

    sheet_value = sheet_name
    if isinstance(sheet_value, str) and sheet_value.strip().isdigit():
        sheet_value = int(sheet_value.strip())
    elif isinstance(sheet_value, str):
        sheet_value = sheet_value.strip() or None

    config: dict[str, Any] = {
        "schema_version": 2,
        "auth": {"app_id": app_id.strip(), "app_secret": None},
        "source": {
            "type": "file",
            "file": {
                "path": source.resolve().as_posix(),
                "sheet_name": sheet_value,
            },
        },
        "target": {
            "type": "bitable",
            "bitable": {
                "app_token": target.app_token,
                "table_id": target.table_id,
                "create_missing_fields": bool(create_missing_fields),
                "api_backend": backend,
                "user_id_type": "open_id",
            },
        },
        "sync": sync,
        "conversion": {
            "strategy": field_strategy,
            "intelligence": {
                "date_confidence": 0.85,
                "choice_confidence": 0.9,
                "boolean_confidence": 0.95,
            },
        },
        "control": {
            "batch_size": int(batch_size),
            "rate_limit_delay": 0.01,
            "max_retries": 3,
            "advanced": {
                "enabled": False,
                "retry": {
                    "strategy": "exponential_backoff",
                    "initial_delay": 0.5,
                    "max_wait_time": None,
                    "multiplier": 2.0,
                    "increment": 0.5,
                },
                "rate_limit": {
                    "strategy": "fixed_wait",
                    "window_size": 1.0,
                    "max_requests": 10,
                },
            },
        },
        "output": {"log_level": "INFO"},
    }
    _validate_config(config)
    return config


def write_config(
    config: Mapping[str, Any], directory: str | Path | None = None
) -> Path:
    folder = Path(directory) if directory else Path(tempfile.mkdtemp(prefix="xtf_gui_"))
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "config.yaml"
    path.write_text(
        yaml.safe_dump(dict(config), allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    return path


def load_config_file(path: str | Path) -> LoadedConfig:
    target_path = Path(path)
    if not target_path.is_file():
        raise GuiError(f"找不到配置文件：{target_path}")
    try:
        document = yaml.safe_load(target_path.read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001
        raise GuiError(f"读取配置失败：{exc}") from exc
    if not isinstance(document, Mapping):
        raise GuiError("配置文件内容不是 YAML 对象。")
    if document.get("schema_version") != 2:
        raise GuiError("只支持 schema_version: 2 的配置。")

    auth = document.get("auth") or {}
    source = document.get("source") or {}
    target = document.get("target") or {}
    sync = document.get("sync") or {}
    conversion = document.get("conversion") or {}
    control = document.get("control") or {}

    target_type = target.get("type")
    ref: TargetRef | None = None
    if target_type == "bitable":
        bitable = target.get("bitable") or {}
        ref = TargetRef(
            kind="bitable",
            app_token=str(bitable.get("app_token") or ""),
            table_id=str(bitable.get("table_id") or ""),
        )
    elif target_type == "sheet":
        sheet = target.get("sheet") or {}
        ref = TargetRef(
            kind="sheet",
            spreadsheet_token=str(sheet.get("spreadsheet_token") or ""),
            sheet_id=str(sheet.get("sheet_id") or "") or None,
        )

    sync_mode = str(sync.get("mode") or "full")
    strategy = sync.get("match_strategy")
    if sync_mode == "clone":
        gui_mode = "clone"
    elif sync_mode == "incremental" and strategy == "append_only":
        gui_mode = "append"
    elif sync_mode == "incremental":
        gui_mode = "incremental"
    elif sync_mode == "overwrite":
        gui_mode = "overwrite"
    else:
        gui_mode = "full"

    index = sync.get("index") or {}
    selective = sync.get("selective") or {}
    selective_columns = None
    if selective.get("enabled") and isinstance(selective.get("columns"), list):
        selective_columns = [str(item) for item in selective["columns"]]

    secret_value = auth.get("app_secret")
    secret = (
        str(secret_value) if isinstance(secret_value, str) and secret_value else None
    )

    source_file = source.get("file") or {}
    sheet_value = source_file.get("sheet_name")
    return LoadedConfig(
        app_id=str(auth.get("app_id") or ""),
        secret=secret,
        source_path=str(source_file.get("path") or ""),
        sheet_name=str(sheet_value) if sheet_value not in (None, "") else None,
        target=ref,
        mode=gui_mode,
        index_column=(
            str(index.get("column")) if index.get("column") not in (None, "") else None
        ),
        selective_columns=selective_columns,
        verify_remote_writes=bool(sync.get("verify_remote_writes", False)),
        create_missing_fields=bool(
            (target.get("bitable") or {}).get("create_missing_fields", True)
        ),
        field_strategy=str(conversion.get("strategy") or "base"),
        backend=str((target.get("bitable") or {}).get("api_backend") or "base_v3"),
        batch_size=int(control.get("batch_size") or 500),
    )


# ---------------------------------------------------------------------------
# CLI execution and result formatting
# ---------------------------------------------------------------------------


def cli_command(
    config_path: str | Path, *, dry_run: bool = False, allow_delete: bool = False
) -> list[str]:
    script = ROOT / "XTF.py"
    if script.is_file():
        import sys

        python = Path(sys.executable)
        if python.name.lower() == "pythonw.exe":
            candidate = python.with_name("python.exe")
            if candidate.is_file():
                python = candidate
        command = [str(python), str(script)]
    else:
        import sys

        candidates = [
            Path(sys.executable).with_name("XTF.exe"),
            Path(sys.executable).with_name("XTF"),
            ROOT / "XTF.exe",
            ROOT / "XTF",
        ]
        executable = next((item for item in candidates if item.is_file()), None)
        if executable is None:
            raise GuiError(
                "找不到 XTF 可执行文件。",
                hint="请从源码目录运行界面，或把 XTF 二进制放在界面程序旁边。",
            )
        command = [str(executable)]
    command.extend(["sync", "--config", str(config_path), "--json"])
    if dry_run:
        command.append("--dry-run")
    if allow_delete:
        command.append("--allow-delete")
    return command


def subprocess_env(secret: str) -> dict[str, str]:
    import os

    env = os.environ.copy()
    env["XTF_APP_SECRET"] = secret
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def parse_cli_json(stdout: str) -> dict[str, Any] | None:
    text = (stdout or "").strip()
    if not text:
        return None
    try:
        value = json.loads(text)
        return dict(value) if isinstance(value, Mapping) else None
    except json.JSONDecodeError:
        pass
    for line in reversed(text.splitlines()):
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, Mapping):
            return dict(value)
    return None


def plan_rows(plan: Mapping[str, Any] | None) -> list[dict[str, Any]]:
    if not plan:
        return []
    rows: list[dict[str, Any]] = []
    for action in plan.get("actions") or []:
        if not isinstance(action, Mapping):
            continue
        kind = str(action.get("kind") or "")
        destructive = bool(action.get("destructive"))
        clears = bool(action.get("clears_values"))
        risk = "危险" if destructive else ("会清空值" if clears else "正常")
        rows.append(
            {
                "kind": ACTION_LABELS.get(kind, kind),
                "raw_kind": kind,
                "count": action.get("count", 0),
                "unit": UNIT_LABELS.get(
                    str(action.get("unit")), str(action.get("unit"))
                ),
                "risk": risk,
                "destructive": destructive,
            }
        )
    return rows


def actions_summary(plan: Mapping[str, Any] | None) -> str:
    rows = plan_rows(plan)
    if not rows:
        return "没有需要执行的写入动作"
    return "；".join(f"{row['kind']} {row['count']} {row['unit']}" for row in rows)


def plan_is_destructive(plan: Mapping[str, Any] | None) -> bool:
    if not plan:
        return False
    if plan.get("destructive"):
        return True
    for action in plan.get("actions") or []:
        if not isinstance(action, Mapping):
            continue
        if action.get("kind") in {"delete_records", "clear_range"}:
            return True
        if action.get("destructive"):
            return True
    return False


def plan_clears_values(plan: Mapping[str, Any] | None) -> bool:
    if not plan:
        return False
    if plan.get("clears_values"):
        return True
    return any(
        bool(action.get("clears_values"))
        for action in plan.get("actions") or []
        if isinstance(action, Mapping)
    )


def exit_label(code: int) -> str:
    return EXIT_LABELS.get(code, f"未知退出码 {code}")


def status_label(status: Any) -> str:
    return STATUS_LABELS.get(str(status or ""), str(status or "未知"))


def payload_error_text(payload: Mapping[str, Any] | None) -> str:
    if not payload:
        return ""
    error = payload.get("error")
    if isinstance(error, Mapping):
        kind = error.get("kind")
        message = error.get("message") or error.get("code") or ""
        return f"{kind}：{message}" if kind else str(message)
    return str(error or "")


def outcome_text(payload: Mapping[str, Any] | None) -> str:
    if not payload:
        return ""
    result = payload.get("result")
    outcome = result.get("outcome") if isinstance(result, Mapping) else None
    if not isinstance(outcome, Mapping):
        outcome = {}
    status = outcome.get("status") or payload.get("status")
    summary = outcome.get("summary") or {}
    accepted = summary.get("accepted_by_unit") or {}
    confirmed = summary.get("confirmed_by_unit") or {}
    parts = [f"状态：{status_label(status)}"]
    if accepted:
        parts.append(
            "接受："
            + "，".join(
                f"{value} {UNIT_LABELS.get(str(key), str(key))}"
                for key, value in accepted.items()
            )
        )
    if confirmed:
        parts.append(
            "确认："
            + "，".join(
                f"{value} {UNIT_LABELS.get(str(key), str(key))}"
                for key, value in confirmed.items()
            )
        )
    if summary.get("confirmation_complete"):
        parts.append("已完成读回确认")
    return "；".join(parts)


# ---------------------------------------------------------------------------
# Small persistence helpers (never store the secret)
# ---------------------------------------------------------------------------


def settings_path() -> Path:
    base = Path(os.environ.get("APPDATA") or Path.home())
    return base / "XTF" / "xtf_gui_settings.json"


def load_settings() -> dict[str, Any]:
    path = settings_path()
    if not path.is_file():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {}
    return dict(value) if isinstance(value, Mapping) else {}


def save_settings(values: Mapping[str, Any]) -> None:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(dict(values), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
