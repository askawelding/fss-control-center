from __future__ import annotations

import sys

import pytest

import tanpa_rumus_testkit as kit
from app.services import tanpa_rumus as tr


@pytest.fixture
def fake_excel(monkeypatch):
    """Ganti xlwings/Excel dengan model palsu (lihat tanpa_rumus_testkit.py) dan pastikan tidak ada proses bocor."""
    kit.reset_fake_excel()
    monkeypatch.setitem(sys.modules, "xlwings", kit.FakeXlwings)
    monkeypatch.setattr(
        tr, "_process_alive", lambda pid: any(app.pid == pid and app.alive for app in kit.FakeApp.registry)
    )
    monkeypatch.setattr(tr, "QUIT_WAIT_POLLS", 2)
    monkeypatch.setattr(tr, "QUIT_POLL_SECONDS", 0.01)
    yield kit.BEHAVIOR
    assert kit.alive_excel_instances() == [], "proses Excel tertinggal"
    kit.reset_fake_excel()
