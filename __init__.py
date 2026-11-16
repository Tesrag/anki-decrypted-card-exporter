# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import json
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from aqt import gui_hooks, mw
from aqt.qt import QAction, QFileDialog, QMessageBox, QProgressDialog, QTimer, Qt
from aqt.browser.previewer import MultiCardPreviewer

_ACTIVE = None
_ADDON_VERSION = "4.0.0"
_RETRY_LIMIT = 5
_PHASE_TIMEOUT_MS = 8000


def _text(v):
    return "" if v is None else str(v).strip()


def _write_xlsx(path: Path, headers, rows):
    def col(n):
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    xml_rows = []
    for ri, row in enumerate([headers] + rows, 1):
        cells = []
        for ci, val in enumerate(row, 1):
            value = escape(_text(val))
            cells.append(
                f'<c r="{col(ci)}{ri}" t="inlineStr"><is>'
                f'<t xml:space="preserve">{value}</t></is></c>'
            )
        xml_rows.append(f'<row r="{ri}">{"".join(cells)}</row>')

    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="题库" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/></Relationships>'
    )
    root = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    )
    types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", types)
        z.writestr("_rels/.rels", root)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def _capture_card_html(html, card, kind):
    """Append a self-contained callback to each preview side."""
    if _ACTIVE is None or not kind.startswith("preview"):
        return html

    event = "question_ready" if kind == "previewQuestion" else "answer_ready"
    card_id = int(card.id)

    script = r"""
<script>
(function() {
  var eventName = __EVENT__;
  var expectedCardId = __CARD_ID__;

  function collectAndSend() {
    var data = {};
    var typeText = (document.querySelector("#type_container") || {}).innerText || "";
    var questionText = (document.querySelector(".q-body") || {}).innerText || "";
    var optionNodes = Array.from(document.querySelectorAll("#q-options .q-answer"));

    if (eventName === "question_ready") {
      data = {
        type: typeText.trim(),
        question: questionText.trim(),
        options: optionNodes.map(function(x) { return x.innerText.trim(); })
      };
      if (!data.question || data.options.length === 0 ||
          data.options.some(function(x) { return !x; })) {
        eventName = "question_incomplete";
      }
    }

    if (eventName === "answer_ready") {
      var articles = Array.from(document.querySelectorAll("article"));
      var answerArticle = articles.find(function(a) {
        var h = a.querySelector(".heading");
        return h && h.innerText.trim() === "答案解析";
      });

      var answerDiv = document.querySelector("#answer_div");
      var answerText = answerDiv ? answerDiv.innerText.trim() : "";
      var articleText = answerArticle ? answerArticle.innerText.trim() : "";

      var letters = [];
      var fullText = answerText + "\n" + articleText;
      var re = /正确答案\s*[:：]\s*([A-ZＡ-Ｚ](?:\s*[,，、和及]?\s*[A-ZＡ-Ｚ])*)/g;
      var match;
      while ((match = re.exec(fullText)) !== null) {
        var found = match[1].match(/[A-ZＡ-Ｚ]/g) || [];
        found.forEach(function(x) { letters.push(x.toUpperCase()); });
        if (letters.length) break;
      }

      var marked = optionNodes.filter(function(x) {
        return x.classList.contains("correct") ||
          !!x.querySelector('img[src*="__correct.svg"], img[src*="correct.svg"]');
      });
      var selectedOptions = marked.length ? marked : optionNodes.filter(function(x) {
        var t = x.innerText.trim();
        var m = t.match(/^([A-ZＡ-Ｚ])\s*[.、．]/);
        return m && letters.indexOf(m[1].toUpperCase()) >= 0;
      });

      var correct = selectedOptions.map(function(x) { return x.innerText.trim(); });
      if (!correct.length && letters.length) correct = letters;

      // Collect the complete explanation region, not only the first sibling.
      // Card templates may split explanations across paragraphs, lists, and wrappers.
      // Preserve the complete rendered text of the entire "答案解析" article.
      // Do not extract only the first sibling after #choice_div: long
      // explanations may be distributed across multiple nested elements.
      var explanation = "";
      if (answerArticle) {
        explanation = (answerArticle.innerText || answerArticle.textContent || "").trim();

        // Remove only the section heading at the beginning and the unrelated
        // statistics section. Keep all remaining paragraphs and line breaks.
        explanation = explanation.replace(/^答案解析\s*/, "");
        explanation = explanation.replace(/\n\s*答题统计[\s\S]*$/, "").trim();

        // If the statistics title is in a separate nested element and the
        // rendered text has no line break before it, cut at the first title.
        var statsIndex = explanation.indexOf("答题统计");
        if (statsIndex >= 0) {
          explanation = explanation.slice(0, statsIndex).trim();
        }
      }
      data = {
        type: typeText,
        question: questionText,
        options: optionNodes.map(function(x) { return x.innerText.trim(); }),
        correctAnswer: correct.join("；"),
        answer: answerText,
        explanation: explanation,
        answerArticleText: articleText
      };
      if (!data.correctAnswer || !data.explanation) {
        eventName = "answer_incomplete";
      }
    }

    pycmd("anki_decrypt_export_v46:" + JSON.stringify({
      event: eventName,
      cardId: expectedCardId,
      data: data
    }));
  }

  // Anki resets onShownHook for each displayed side. Register only for this
  // render; retry registration briefly if the page has not initialized it yet.
  var attempts = 0;
  function register() {
    attempts += 1;
    if (Array.isArray(window.onShownHook)) {
      window.onShownHook.push(function() {
        collectAndSend();
      });
    } else if (attempts < 40) {
      setTimeout(register, 50);
    } else {
      pycmd("anki_decrypt_export_v46:" + JSON.stringify({
        event: "hook_unavailable",
        cardId: expectedCardId,
        data: {}
      }));
    }
  }
  register();
})();
</script>
"""
    script = script.replace("__EVENT__", json.dumps(event))
    script = script.replace("__CARD_ID__", str(card_id))
    return html + script


class ExportPreviewer(MultiCardPreviewer):
    def __init__(self, browser, cards, output_dir):
        self.cards = cards
        self.index = 0
        self.output_dir = Path(output_dir)
        self.records = []
        self.failures = []
        self.cancelled = False
        self.phase = "question"
        self.retry_count = 0
        self._web_ready = False
        self._accept_callbacks = False
        self._startup_kick_done = False
        self._partial_data = {}
        self.phase_timer = QTimer()
        self.phase_timer.setSingleShot(True)
        self.phase_timer.timeout.connect(self._on_phase_timeout)
        self.progress = QProgressDialog(
            "正在等待 Anki 渲染卡片……", "取消", 0, len(cards), browser
        )
        self.progress.setWindowTitle("批量提取解密题库 V4.6")
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress.setMinimumDuration(0)
        self.progress.canceled.connect(self.cancel)
        super().__init__(parent=browser, mw=mw, on_close=self._closed)

    def card(self):
        if 0 <= self.index < len(self.cards):
            return self.cards[self.index]
        return None

    def card_changed(self):
        # Anki resets the previewer to the question side only when the card
        # identity changes. Returning True on every render resets the answer
        # side and prevents the exporter from ever completing a card.
        current = self.card()
        if current is None:
            return True
        current_id = int(current.id)
        changed = current_id != getattr(self, "_last_card_id", None)
        self._last_card_id = current_id
        return changed

    def start(self):
        self.progress.show()
        self._set_progress()
        # open() creates the native preview webview. The initial render may be
        # queued before reviewer.js has finished loading, so V4.6 waits for
        # loadFinished and then explicitly renders the first selected card.
        self.open()

    def _setup_web_view(self):
        super()._setup_web_view()
        assert self._web is not None
        self._web.loadFinished.connect(self._on_web_load_finished)

    def _on_web_load_finished(self, ok):
        if self.cancelled or self._startup_kick_done:
            return
        if not ok:
            self.phase_timer.start(1500)
            return
        self._web_ready = True
        # Allow reviewer.js and its onShownHook globals to initialize.
        QTimer.singleShot(1200, self._start_first_render)

    def _start_first_render(self):
        if self.cancelled or self._startup_kick_done:
            return
        self._startup_kick_done = True
        self._accept_callbacks = True
        self.phase = "question"
        self._state = "question"
        self._last_state = None
        self._card_changed = True
        self._set_progress()
        self.render_card()
        self._arm_timeout()

    def _set_progress(self):
        total = len(self.cards)
        self.progress.setLabelText(
            f"第 {min(self.index + 1, total)}/{total} 张："
            f"{'正面' if self.phase == 'question' else '背面'}"
        )

    def _arm_timeout(self):
        self.phase_timer.stop()
        self.phase_timer.start(_PHASE_TIMEOUT_MS)

    def _on_phase_timeout(self):
        if self.cancelled or self.index >= len(self.cards):
            return
        if self.retry_count < _RETRY_LIMIT:
            self.retry_count += 1
            self._force_rerender()
        else:
            c = self.card()
            cid = str(c.id) if c else ""
            self.failures.append({
                "card_id": cid,
                "phase": self.phase,
                "reason": "等待渲染回传超时，重试次数已用尽"
            })
            if self.phase == "question":
                current = self.card()
                failed = {
                    "card_id": str(current.id) if current else "",
                    "note_id": str(current.nid) if current else "",
                    "type": "", "question": "", "options": [],
                    "correctAnswer": "", "answer": "", "explanation": "",
                    "capture_error": "正面渲染或回传超时"
                }
                self.records.append(failed)
                self.progress.setValue(len(self.records))
                self._on_next_card()
            else:
                # Keep a placeholder row for this Card ID instead of silently
                # dropping it from the export.
                partial = dict(self._partial_data)
                current = self.card()
                if current is not None:
                    partial["card_id"] = str(current.id)
                    partial["note_id"] = str(current.nid)
                partial.setdefault("type", "")
                partial.setdefault("question", "")
                partial.setdefault("options", [])
                partial.setdefault("correctAnswer", "")
                partial.setdefault("answer", "")
                partial.setdefault("explanation", "")
                partial["capture_error"] = "背面渲染或回传超时"
                self.records.append(partial)
                self.progress.setValue(len(self.records))
                self._on_next_card()

    def _force_rerender(self):
        self._last_state = None
        self._card_changed = True
        self._set_progress()
        self.render_card()
        self._arm_timeout()

    def _on_bridge_cmd(self, cmd):
        if not cmd.startswith("anki_decrypt_export_v46:"):
            return super()._on_bridge_cmd(cmd)

        try:
            payload = json.loads(cmd.split(":", 1)[1])
        except Exception:
            return

        if self.cancelled or not self._accept_callbacks or self.index >= len(self.cards):
            return

        current = self.card()
        if current is None:
            return

        # Ignore stale callbacks from a previous card or side.
        if str(payload.get("cardId", "")) != str(current.id):
            return

        event = payload.get("event")
        expected = "question_ready" if self.phase == "question" else "answer_ready"

        if event == "hook_unavailable":
            self._on_phase_timeout()
            return
        if event in ("question_incomplete", "answer_incomplete"):
            if self.retry_count < _RETRY_LIMIT:
                self.retry_count += 1
                self._force_rerender()
            else:
                self._on_phase_timeout()
            return
        if event != expected:
            return

        self.phase_timer.stop()
        self.retry_count = 0

        if event == "question_ready":
            # Keep the front-side snapshot as a fallback if answer rendering
            # fails. It also proves this exact card has completed its front.
            self._partial_data = payload.get("data") or {}
            self.phase = "answer"
            self._set_progress()
            self._on_next()
            self._arm_timeout()
            return

        data = payload.get("data") or {}
        data["card_id"] = str(current.id)
        data["note_id"] = str(current.nid)
        self.records.append(data)
        self.progress.setValue(len(self.records))

        # In MultiCardPreviewer, _on_next() while on the answer side calls
        # _on_next_card(). Advancing here keeps navigation fully automatic.
        self._on_next()
        if self.index < len(self.cards):
            self._arm_timeout()

    def _on_next_card(self):
        self.index += 1
        if self.index >= len(self.cards):
            self.finish()
            return

        self.phase = "question"
        self.retry_count = 0
        self._partial_data = {}
        self._state = "question"
        self._last_state = None
        self._card_changed = True
        self._set_progress()
        self.render_card()
        self._arm_timeout()

    def _on_prev_card(self):
        pass

    def _updateButtons(self):
        if getattr(self, "_prev", None):
            self._prev.setEnabled(False)
        if getattr(self, "_next", None):
            self._next.setEnabled(False)

    def cancel(self):
        self.cancelled = True
        self.phase_timer.stop()
        self.progress.close()
        QMessageBox.information(
            self, "已停止",
            f"已成功读取 {len(self.records)} 张。"
            f"异常阶段 {len(self.failures)} 项。"
            "未执行复习评分或修改笔记。"
        )
        self.close()

    def _closed(self):
        self.phase_timer.stop()
        self.progress.close()
        global _ACTIVE
        _ACTIVE = None

    def finish(self):
        self.phase_timer.stop()
        headers = [
            "card_id", "note_id", "type", "question", "options",
            "correctAnswer", "answer", "explanation", "status"
        ]
        rows = []
        failed_ids = {x["card_id"] for x in self.failures}
        for d in self.records:
            options = d.get("options") or []
            ans = _text(d.get("correctAnswer"))
            expl = _text(d.get("explanation"))
            question = _text(d.get("question"))
            capture_error = _text(d.get("capture_error"))
            status = ("采集失败：" + capture_error) if capture_error else (
                "OK" if ans and expl and question else (
                    "题目缺失" if not question else (
                        "答案缺失" if not ans else "解析缺失"
                    )
                )
            )
            rows.append([
                d.get("card_id", ""), d.get("note_id", ""),
                d.get("type", ""), question,
                "\n".join(options), ans, d.get("answer", ""),
                expl or d.get("answerArticleText", ""), status
            ])

        stem = "江苏省实验动物专业技能培训记录卡_解密题库_v4.6"
        csv_path = self.output_dir / f"{stem}.csv"
        xlsx_path = self.output_dir / f"{stem}.xlsx"
        error_path = self.output_dir / f"{stem}_异常记录.csv"

        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)

        _write_xlsx(xlsx_path, headers, rows)

        with open(error_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=["card_id", "phase", "reason"])
            writer.writeheader()
            writer.writerows(self.failures)

        failed_rows = sum(1 for r in rows if r[-1] != "OK")
        QMessageBox.information(
            self, "V4.6 提取完成",
            f"计划处理：{len(self.cards)} 张\n"
            f"成功回传：{len(self.records)} 张\n"
            f"字段异常：{failed_rows} 张\n"
            f"渲染超时异常：{len(self.failures)} 项\n\n"
            f"CSV：{csv_path}\nExcel：{xlsx_path}\n异常记录：{error_path}\n\n"
            "本插件不触发复习评分，不修改笔记字段。"
        )
        self.close()


def _selected_ids(browser):
    method = getattr(browser, "selectedCards", None)
    if not callable(method):
        return []
    try:
        return list(method())
    except Exception:
        return []


def _start(browser):
    global _ACTIVE
    if _ACTIVE is not None:
        QMessageBox.warning(browser, "任务运行中", "已有提取任务正在运行。")
        return

    ids = _selected_ids(browser)
    if not ids:
        QMessageBox.warning(
            browser, "没有选中卡片",
            "请在浏览器列表中选中需要导出的卡片。"
        )
        return

    folder = QFileDialog.getExistingDirectory(
        browser, "选择导出目录", str(Path.home() / "Desktop")
    )
    if not folder:
        return

    try:
        cards = [mw.col.get_card(cid) for cid in ids]
    except Exception as exc:
        QMessageBox.critical(browser, "读取卡片失败", str(exc))
        return

    _ACTIVE = ExportPreviewer(browser, cards, folder)
    _ACTIVE.start()


def _browser_menu(browser):
    action = QAction("批量导出解密题库 V4.6（CSV + Excel）", browser)
    action.triggered.connect(lambda _=False, b=browser: _start(b))
    browser.form.menuEdit.addSeparator()
    browser.form.menuEdit.addAction(action)


gui_hooks.card_will_show.append(_capture_card_html)
gui_hooks.browser_menus_did_init.append(_browser_menu)
