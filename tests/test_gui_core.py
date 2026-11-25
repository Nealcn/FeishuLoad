"""Tk-free tests for the optional desktop GUI helpers."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import yaml

from api.bitable_backend import FieldKind, FieldSchema
from gui import core
from xtf_cli.config import validate_v2_document


def make_source(tmp_path: Path) -> core.SourceInfo:
    path = tmp_path / "data.xlsx"
    pd.DataFrame(
        {
            "ID": ["1", "2"],
            "Name": ["Alice", "Bob"],
            "Amount": [10, 20],
        }
    ).to_excel(path, index=False)
    return core.read_source(path)


def make_target(records=None, writable=True) -> core.TargetInfo:
    schema = FieldSchema(
        id="fld_id",
        name="ID",
        kind=FieldKind.TEXT,
        multiple=False,
        writable=writable,
        raw_type=1,
    )
    field = core.TargetField(
        name="ID",
        kind="text",
        writable=writable,
        multiple=False,
        raw_type=1,
        schema=schema,
    )
    return core.TargetInfo(
        ref=core.TargetRef(
            kind="bitable",
            app_token="app_demo",
            table_id="tbl_demo",
        ),
        table_name="Demo",
        fields=[field],
        records=records or [{"ID": "1"}, {"ID": "2"}],
        backend="base_v3",
    )


def test_parse_base_url_extracts_tokens():
    ref = core.parse_target_url(
        "https://hisenseex.feishu.cn/base/app_demo?table=tbl_demo&view=vew_demo"
    )
    assert ref.kind == "bitable"
    assert ref.app_token == "app_demo"
    assert ref.table_id == "tbl_demo"
    assert ref.view_id == "vew_demo"


def test_parse_sheet_url_is_recognized_but_not_supported_by_gui():
    ref = core.parse_target_url("https://example.feishu.cn/sheets/sht_demo?sheet=abc")
    assert ref.kind == "sheet"
    assert ref.spreadsheet_token == "sht_demo"
    assert ref.sheet_id == "abc"


def test_read_source_rejects_duplicate_headers(tmp_path):
    path = tmp_path / "bad.csv"
    path.write_text("ID,ID\n1,2\n", encoding="utf-8")
    with pytest.raises(core.GuiError):
        core.read_source(path)


def test_read_source_returns_preview_and_stats(tmp_path):
    info = make_source(tmp_path)
    assert info.rows == 2
    assert info.columns == ["ID", "Name", "Amount"]
    assert info.stats["ID"].nonempty == 2
    assert info.stats["ID"].unique == 2
    assert info.preview[0][:2] == ["1", "Alice"]


def test_sanitize_source_keeps_only_requested_columns(tmp_path):
    info = make_source(tmp_path)
    output = core.sanitize_source(info, ["ID", "Name"], tmp_path / "filtered.xlsx")
    cleaned = pd.read_excel(output, dtype=object)
    assert list(cleaned.columns) == ["ID", "Name"]
    assert len(cleaned) == 2


def test_build_config_append_omits_index_and_secret(tmp_path):
    info = make_source(tmp_path)
    ref = core.TargetRef(kind="bitable", app_token="app_demo", table_id="tbl_demo")
    config = core.build_config(
        app_id="cli_demo",
        source_path=info.path,
        target=ref,
        mode="append",
    )
    validate_v2_document(config)
    assert config["auth"]["app_secret"] is None
    assert config["sync"]["mode"] == "incremental"
    assert config["sync"]["match_strategy"] == "append_only"
    assert "index" not in config["sync"]
    assert config["sync"]["selective"]["enabled"] is False


def test_build_config_clone_omits_match_strategy_and_index(tmp_path):
    info = make_source(tmp_path)
    ref = core.TargetRef(kind="bitable", app_token="app_demo", table_id="tbl_demo")
    config = core.build_config(
        app_id="cli_demo",
        source_path=info.path,
        target=ref,
        mode="clone",
    )
    validate_v2_document(config)
    assert config["sync"]["mode"] == "clone"
    assert "match_strategy" not in config["sync"]
    assert "index" not in config["sync"]


def test_build_config_full_selective_includes_index(tmp_path):
    info = make_source(tmp_path)
    ref = core.TargetRef(kind="bitable", app_token="app_demo", table_id="tbl_demo")
    config = core.build_config(
        app_id="cli_demo",
        source_path=info.path,
        target=ref,
        mode="full",
        index_column="ID",
        selective_columns=["Name"],
    )
    validate_v2_document(config)
    assert config["sync"]["mode"] == "full"
    assert config["sync"]["match_strategy"] == "by_key"
    assert config["sync"]["index"]["column"] == "ID"
    assert config["sync"]["selective"]["enabled"] is True
    assert config["sync"]["selective"]["columns"] == ["Name", "ID"]


@pytest.mark.parametrize("mode", ["incremental", "full", "overwrite"])
def test_build_config_key_modes_require_index(tmp_path, mode):
    info = make_source(tmp_path)
    ref = core.TargetRef(kind="bitable", app_token="app_demo", table_id="tbl_demo")
    with pytest.raises(core.GuiError):
        core.build_config(
            app_id="cli_demo",
            source_path=info.path,
            target=ref,
            mode=mode,
        )


def test_build_config_rejects_secret_in_yaml(tmp_path):
    info = make_source(tmp_path)
    ref = core.TargetRef(kind="bitable", app_token="app_demo", table_id="tbl_demo")
    config = core.build_config(
        app_id="cli_demo",
        source_path=info.path,
        target=ref,
        mode="append",
    )
    text = yaml.safe_dump(config, allow_unicode=True)
    assert "app_secret: null" in text


def test_index_candidates_marks_safe_unique_columns(tmp_path):
    source = make_source(tmp_path)
    target = make_target()
    candidates = core.index_candidates(source, target)
    assert candidates
    assert candidates[0].name == "ID"
    assert candidates[0].safe is True


def test_index_candidates_rejects_target_duplicates(tmp_path):
    source = make_source(tmp_path)
    target = make_target(records=[{"ID": "1"}, {"ID": "1"}])
    candidates = core.index_candidates(source, target)
    assert candidates
    assert candidates[0].safe is False
    assert "重复" in candidates[0].reason


def test_index_candidates_skips_read_only_fields(tmp_path):
    source = make_source(tmp_path)
    target = make_target(writable=False)
    assert core.index_candidates(source, target) == []


def test_plan_rows_and_destructive_detection():
    plan = {
        "actions": [
            {
                "kind": "delete_records",
                "count": 3,
                "unit": "record",
                "destructive": True,
                "clears_values": False,
            }
        ],
        "warnings": ["careful"],
    }
    rows = core.plan_rows(plan)
    assert rows[0]["kind"] == "删除记录"
    assert rows[0]["risk"] == "危险"
    assert core.plan_is_destructive(plan) is True
    assert core.plan_is_destructive({"actions": []}) is False


def test_load_config_file_round_trip(tmp_path):
    info = make_source(tmp_path)
    ref = core.TargetRef(kind="bitable", app_token="app_demo", table_id="tbl_demo")
    config = core.build_config(
        app_id="cli_demo",
        source_path=info.path,
        target=ref,
        mode="full",
        index_column="ID",
        selective_columns=["Name"],
    )
    config["auth"]["app_secret"] = "test_secret"
    path = tmp_path / "config.yaml"
    path.write_text(
        yaml.safe_dump(config, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    loaded = core.load_config_file(path)
    assert loaded.app_id == "cli_demo"
    assert loaded.secret == "test_secret"
    assert loaded.target is not None
    assert loaded.target.app_token == "app_demo"
    assert loaded.mode == "full"
    assert loaded.index_column == "ID"
    assert loaded.selective_columns == ["Name", "ID"]


def test_cli_command_uses_source_entrypoint():
    command = core.cli_command(Path("config.yaml"), dry_run=True)
    assert command[0].endswith(("python.exe", "python"))
    assert command[1].endswith("XTF.py")
    assert "--dry-run" in command
    assert "--json" in command
