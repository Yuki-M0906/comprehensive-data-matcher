# -*- coding: utf-8 -*-
"""Web 版（app.py）を画面なしで動かすテスト（streamlit.testing）。"""

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from app_text import TEXT

APP = str(Path(__file__).resolve().parent.parent / "app.py")
TIMEOUT = 120

# 古い streamlit.testing（Python 3.9 で入る Streamlit 1.50 など）は、2 回目以降の実行で
# segmented_control の値を正しく渡せない（実際の画面では問題ない）。その版では、操作のテストを飛ばす
NEW_APPTEST = hasattr(AppTest, "segmented_control")
needs_new_apptest = pytest.mark.skipif(
    not NEW_APPTEST, reason="この版の streamlit.testing は segmented_control を扱えない")


def run_app(lang=None):
    at = AppTest.from_file(APP, default_timeout=TIMEOUT)
    if lang:
        at.query_params["lang"] = lang
    return at.run()


def metrics(at):
    return {m.label: m.value for m in at.metric}


def button(at, label):
    return next(b for b in at.button if b.label == label)


def test_text_keys_are_the_same_in_both_languages():
    assert set(TEXT["ja"]) == set(TEXT["en"])


def test_sample_results_are_shown_on_first_load():
    at = run_app()
    assert not at.exception
    assert at.title[0].value == TEXT["ja"]["title"]
    m = metrics(at)
    assert m[TEXT["ja"]["m_rows"]] == "34"
    assert m[TEXT["ja"]["m_ok"]].startswith("19")
    assert m[TEXT["ja"]["m_review"]] == "15"


def test_english_by_url():
    at = run_app("en")
    assert not at.exception
    assert at.title[0].value == TEXT["en"]["title"]
    assert metrics(at)[TEXT["en"]["m_rows"]] == "34"


@needs_new_apptest
def test_settings_change_the_results():
    at = run_app()
    at.slider(key="threshold__ja").set_value(0.9).run()
    assert not at.exception
    assert metrics(at)[TEXT["ja"]["m_review"]] == "33"
    button(at, TEXT["ja"]["reset"]).click().run()
    assert at.slider(key="threshold__ja").value == pytest.approx(0.5)
    assert metrics(at)[TEXT["ja"]["m_review"]] == "15"


@needs_new_apptest
def test_company_sample():
    at = run_app()
    at.segmented_control(key="sample__ja").set_value("companies").run()
    assert not at.exception
    m = metrics(at)
    assert m[TEXT["ja"]["m_rows"]] == "24"
    assert m[TEXT["ja"]["m_review"]] == "0"


@needs_new_apptest
def test_paste_flow():
    at = run_app()
    at.text_area(key="paste_v__ja").set_value("ＡＢＣ 株式会社\n\nXYZ\n")
    at.text_area(key="paste_m__ja").set_value("ABC 株式会社\nXYZ 商事\n")
    at.run()
    button(at, TEXT["ja"]["run_paste"]).click().run()
    assert not at.exception
    assert metrics(at)[TEXT["ja"]["m_rows"]] == "2"  # 空の行は数えない
    assert any("貼り付けた内容" in c.value for c in at.caption)


@needs_new_apptest
def test_language_switch_keeps_settings():
    at = run_app()
    at.slider(key="w_lev__ja").set_value(40).run()
    at.segmented_control(key="lang_pick").set_value("English").run()
    assert not at.exception
    assert at.title[0].value == TEXT["en"]["title"]
    assert at.slider(key="w_lev__en").value == 40
    at.segmented_control(key="lang_pick").set_value("日本語").run()
    assert at.slider(key="w_lev__ja").value == 40


@needs_new_apptest
def test_try_two_names():
    at = run_app()
    at.text_input(key="try_a__ja").set_value("Ｔシャツ 白").run()
    at.text_input(key="try_b__ja").set_value("Tシャツ 白").run()
    assert not at.exception
    assert metrics(at)[TEXT["ja"]["try_total"]] == "1.000"
