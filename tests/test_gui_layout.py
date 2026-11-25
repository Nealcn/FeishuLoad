"""Layout regression tests for the XTF desktop GUI.

These tests are skipped when no display is available so the default CI path
stays headless-safe.  They assert that long content never pushes action
buttons out of the visible area.
"""

from __future__ import annotations

import os
import sys

import pytest

pytest.importorskip("tkinter")

from tkinter import ttk  # noqa: E402

from gui import core  # noqa: E402
from gui.app import XTFGui  # noqa: E402


def _display_available() -> bool:
    if sys.platform.startswith("linux") and not os.environ.get("DISPLAY"):
        return False
    return True


pytestmark = pytest.mark.skipif(
    not _display_available(), reason="no display available for Tk layout tests"
)


@pytest.fixture(scope="module")
def app():
    """One shared window for the module.

    Tk cannot rebuild its root interpreter inside one process, so the
    window is created once and reused.  Each test resets the state it
    changes.
    """
    window = XTFGui()
    window.geometry("1160x820")
    window.update_idletasks()
    window.update()
    try:
        yield window
    finally:
        try:
            window.destroy()
        except Exception:  # noqa: BLE001
            pass


@pytest.fixture(autouse=True)
def _reset_state(app):
    """Restore the shared window to a known state before each test."""
    app.mode_var.set("append")
    app.backend_var.set("base_v3")
    app.show_dangerous_var.set(False)
    app._toggle_dangerous()
    app.mode_description_var.set("")
    app.mode_summary_var.set("")
    app.mode_risk_var.set("")
    app.index_hint.configure(text="读取源文件和目标表后自动检查。")
    app.update_idletasks()
    app.update()
    yield


def _tab(window: XTFGui, index: int):
    return window.notebook.nametowidget(window.notebook.tabs()[index])


def _settings_action_row(window: XTFGui):
    tab = _tab(window, 2)
    rows = sorted(
        child for child in tab.winfo_children() if child.grid_info().get("row") == 1
    )
    assert rows, "settings tab must keep an action row outside the scroll panel"
    return rows[0]


def test_settings_tab_action_row_is_inside_the_visible_area(app):
    """The 生成预览 button must never be clipped by the window edge."""
    tab = _tab(app, 2)
    action_row = _settings_action_row(app)
    app.update_idletasks()
    app.update()
    bottom = action_row.winfo_y() + action_row.winfo_height()
    assert bottom <= tab.winfo_height(), (
        f"settings action row ends at {bottom} but the tab is only "
        f"{tab.winfo_height()} tall"
    )


def test_settings_tab_content_scrolls_instead_of_overflowing(app):
    """Long mode summaries must stay reachable through the scroll panel."""
    tab = _tab(app, 2)
    panel = next(
        child for child in tab.winfo_children() if child.grid_info().get("row") == 0
    )
    app.update_idletasks()
    app.update()
    canvas = panel.winfo_children()[0]
    region = canvas.cget("scrollregion")
    assert region, "scroll panel must publish a scroll region"
    bottom = float(region.split()[3])
    assert (
        bottom >= canvas.winfo_height()
    ), "the scroll panel content must be at least as tall as the viewport"


def test_settings_panel_inner_width_follows_the_canvas(app):
    """Content must use the full width instead of collapsing to its minimum."""
    tab = _tab(app, 2)
    panel = next(
        child for child in tab.winfo_children() if child.grid_info().get("row") == 0
    )
    app.update_idletasks()
    app.update()
    canvas = panel.winfo_children()[0]
    assert (
        panel.inner.winfo_width() >= canvas.winfo_width() - 2
    ), "inner frame width must track the canvas width"


def test_long_index_hint_wraps_instead_of_overflowing(app):
    """A long single-line hint must not exceed the window width."""
    app.notebook.select(2)
    app.index_hint.configure(
        text="没有唯一且非空的列，只能使用“只追加”。"
        "交货单：源列存在重复值；订单号：源列存在重复值；"
        "订单价税合计：源列存在重复值；型号：源列存在重复值；"
        "商家：源列存在重复值；物料编码：源列存在重复值"
    )
    app.update_idletasks()
    app.update()
    available = app.winfo_width()
    assert (
        app.index_hint.winfo_reqwidth() <= available
    ), "index hint must wrap inside the window width"


def test_mode_description_and_risk_are_separate_variables(app):
    """Facts and risk notes must not share one overflowing label."""
    app.mode_var.set("append")
    app._update_mode_description()
    app._update_mode_summary()
    assert app.mode_description_var.get()
    assert app.mode_risk_var.get() == ""
    app.mode_var.set("overwrite")
    app._update_mode_description()
    app._update_mode_summary()
    assert "会删除" in app.mode_risk_var.get()
    assert "会删除" not in app.mode_summary_var.get()


def test_mode_summary_stays_short_enough_for_one_row(app):
    """The factual summary line must not grow with the number of columns."""
    app.mode_var.set("overwrite")
    app._update_mode_summary()
    app.update_idletasks()
    app.update()
    assert len(app.mode_summary_var.get()) <= 120


def test_backend_choice_does_not_change_the_settings_layout(app):
    """base_v3 and bitable_v1 must lay out identically."""
    heights = {}
    for backend in ("base_v3", "bitable_v1"):
        app.backend_var.set(backend)
        app.mode_var.set("append")
        app._update_mode_summary()
        app.update_idletasks()
        app.update()
        tab = _tab(app, 2)
        heights[backend] = tuple(
            child.winfo_height()
            for child in sorted(
                tab.winfo_children(),
                key=lambda item: item.grid_info().get("row") or 0,
            )
        )
    assert heights["base_v3"] == heights["bitable_v1"]


def test_core_mode_labels_cover_every_mode():
    for key in ("append", "incremental", "full", "overwrite", "clone"):
        assert key in core.MODE_LABELS


def test_table_dropdown_lives_on_its_own_row(app):
    """The data table picker must be a dropdown row under the Base link."""
    tab = _tab(app, 0)
    rows: dict[int, list] = {}
    for child in tab.winfo_children():
        if child.winfo_class().lower() != "tlabelframe":
            continue
        for inner in child.winfo_children():
            row = inner.grid_info().get("row")
            if row is not None:
                rows.setdefault(row, []).append(inner)
    table_row = None
    for row, widgets in rows.items():
        if any(widget is app.table_combo for widget in widgets):
            table_row = row
            break
    assert table_row is not None, "the table dropdown must be gridded"
    # Row 0 holds the Base link entry; the dropdown must be below it.
    assert table_row > 0, "the table dropdown must sit below the Base link"
    assert any(
        widget.winfo_class() == "TEntry" for widget in rows.get(0, [])
    ), "the Base link entry must stay on the first row"


def test_table_dropdown_uses_a_combobox_not_a_dialog(app):
    """Selecting a table must not require opening a separate window."""
    assert isinstance(app.table_combo, ttk.Combobox)
    assert str(app.table_combo.cget("state")) == "readonly"


def test_table_selection_updates_the_link_and_label(app):
    """Picking a table fills the ?table= parameter and the parsed label."""
    app._table_items = {"tbl_demo": {"table_id": "tbl_demo", "name": "Demo"}}
    app._table_labels = {"tbl_demo": "Demo  (tbl_demo)"}
    app.table_combo.configure(values=["Demo  (tbl_demo)"])
    app._target_ref = core.TargetRef(kind="bitable", app_token="app_demo")
    app.table_combo.set("Demo  (tbl_demo)")
    assert app._apply_table_selection("tbl_demo") is True
    assert app.target_url_var.get() == (
        "https://feishu.cn/base/app_demo?table=tbl_demo"
    )
    assert "Demo" in app.parsed_target_var.get()


def test_changing_the_link_clears_the_stale_table_list(app):
    """A new Base link must not keep the previous Base's table list."""
    app._table_items = {"tbl_old": {"table_id": "tbl_old", "name": "Old"}}
    app._table_labels = {"tbl_old": "Old  (tbl_old)"}
    app.table_combo.configure(values=["Old  (tbl_old)"])
    app.table_combo.set("Old  (tbl_old)")
    app.target_url_var.set("https://feishu.cn/base/app_new")
    app.update_idletasks()
    assert tuple(app.table_combo.cget("values")) == ()
    assert app._table_items == {}


def test_programmatic_url_update_does_not_clear_the_table_list(app):
    """Selecting a table rewrites the URL and must not wipe the list."""
    app._table_items = {"tbl_demo": {"table_id": "tbl_demo", "name": "Demo"}}
    app._table_labels = {"tbl_demo": "Demo  (tbl_demo)"}
    app.table_combo.configure(values=["Demo  (tbl_demo)"])
    app._target_ref = core.TargetRef(kind="bitable", app_token="app_demo")
    app._apply_table_selection("tbl_demo")
    assert app._table_items, "the table list must survive a programmatic update"
    assert tuple(app.table_combo.cget("values")) == ("Demo  (tbl_demo)",)
