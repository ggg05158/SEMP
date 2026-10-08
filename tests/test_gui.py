"""Integration checks against real Qt widgets, rendered without opening test windows."""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from PySide6.QtCore import Qt, QMimeData, QPointF, QPoint, QSettings, QDate
from PySide6.QtGui import QDropEvent, QDragEnterEvent, QDragMoveEvent, QTextCursor, QFontDatabase, QFont, QColor
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox, QInputDialog, QPushButton, QTableView, QComboBox, QLineEdit, QTextEdit, QLabel
from SEMP import MainWindow
from SEMP import Store
from SEMP import STYLE
from SEMP import EntryDialog, ImportDialog, NodeDialog, CollectionDialog, PurposeChangeDialog
from SEMP import ImportRow, read_import
from SEMP import NODE_MIME, IMPORT_MIME, TagEdit, node_caption, ColorPaletteDialog, PurposeButtons, NoteCard

APP = QApplication.instance() or QApplication([])
# Qt's offscreen platform does not discover the Windows system font database.
for font in ("malgun.ttf", "malgunbd.ttf", "segoeui.ttf"):
    path = Path("C:/Windows/Fonts") / font
    if path.exists(): QFontDatabase.addApplicationFont(str(path))
APP.setFont(QFont("Malgun Gothic",10))
APP.setStyle("Fusion")
APP.setStyleSheet(STYLE)


class GuiTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        project = Store.create(self.tmp.name, "체계 장비 관리", default_groups=False)
        self.group = project.add_node("group", "장비")
        self.device = project.add_node("device", "다기능 통신장치", self.group, purpose="현장 통신", description="현장 시험 및 통신 상태 확인용 장치", tag_text="#통신 #현장A")
        self.units = [project.add_node("unit", "다기능 통신장치", self.device, serial=f"COM-{i:03}", purpose="현장 통신", tag_text="#현장A") for i in range(1, 7)]
        self.project_file = project.path
        project.close()
        self.settings = QSettings(str(Path(self.tmp.name) / "settings.ini"), QSettings.Format.IniFormat)
        self.window = MainWindow(settings=self.settings)
        def fail_on_error(exc):
            raise exc
        self.window.show_error = fail_on_error
        self.window.load_project(self.project_file)
        self.window.show()
        APP.processEvents()

    def tearDown(self):
        self.window.close()
        APP.processEvents()
        self.tmp.cleanup()

    def test_record_period_targets_tabs_and_search_scope(self):
        w = self.window
        first = w.store.save_entry(title="기간 점검", category="점검", day="2026-09-29", end_day="2026-10-02", html="기간 점검", plain="기간 점검", event=True, node_ids=self.units[:3])
        second = w.store.save_entry(title="다른 점검", category="점검", day="2026-09-29", html="다른 점검", plain="다른 점검", event=True, node_ids=self.units[3:])
        w.refresh()
        w.calendar.setSelectedDate(QDate(2026, 9, 30))
        APP.processEvents()
        self.assertEqual([e["id"] for e in w.calendar.events_by_day["2026-09-30"]], [first])
        text = w.calendar_entries.item(0).text()
        self.assertIn("COM-001 외 2개", text)
        self.assertIn("2026-09-29 ~ 2026-10-02", text)
        w.calendar.setSelectedDate(QDate(2026, 10, 2))
        APP.processEvents()
        self.assertEqual(len(w.calendar.events_by_day["2026-10-02"]), 1)
        self.assertNotIn("2026-10-03", w.calendar.events_by_day)
        w.select_node(self.units[0])
        node_tab = w.view_tabs.currentIndex()
        w.search.setText("기간 점검")
        w.set_mode("search")
        result = next(w.search_results.topLevelItem(i) for i in range(w.search_results.topLevelItemCount()) if w.search_results.topLevelItem(i).data(0, Qt.ItemDataRole.UserRole) == ("entry", first))
        w.activate_search(result, 0)
        record_tab = w.view_tabs.currentIndex()
        self.assertEqual(w.events.count(), 1)
        self.assertEqual(w.event_id, first)
        self.assertNotEqual(node_tab, record_tab)
        w.show_home()
        self.assertIsNone(w.selected_id)
        w.view_tabs.setCurrentIndex(node_tab)
        self.assertEqual(w.selected_id, self.units[0])
        w.view_tabs.setCurrentIndex(record_tab)
        self.assertEqual(w.event_id, first)
        self.assertEqual(w.events.count(), 1)
        def save(dialog):
            dialog.rich.editor.setPlainText("검색과 무관한 새 기록")
            return QDialog.DialogCode.Accepted
        with patch.object(EntryDialog, "exec", save): w.new_event()
        self.assertEqual(w.mode, "events")
        self.assertEqual(w.event_id, first)
        self.assertEqual(w.view_tabs.currentIndex(), record_tab)
        self.assertEqual(w.events.count(), 1)
        APP.processEvents()
        w.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "record-tabs.png"))
        w.open_record(second)
        self.assertTrue(w.events.isHidden())
        self.assertEqual(w.event_id, second)
        dialog = EntryDialog(w.store, event=True, entry=w.store.entry(first), parent=w)
        self.assertEqual(dialog.end_day.date(), QDate(2026, 10, 2))
        dialog.show()
        APP.processEvents()
        dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "record-period.png"))
        QTest.mouseClick(dialog.end_calendar_button, Qt.MouseButton.LeftButton)
        self.assertTrue(dialog.end_mini_calendar.isVisible())
        dialog.end_mini_calendar.clicked.emit(QDate(2026, 10, 4))
        self.assertEqual(dialog.values()["end_day"], "2026-10-04")
        dialog.reject()
        with self.assertRaises(ValueError):
            w.store.save_entry(title="역순", category="기타", day="2026-10-02", end_day="2026-09-30", html="x", plain="x", event=True)
        self.assertTrue(w.close_all_tabs_btn.isEnabled())
        w.close_all_views()
        self.assertEqual(w.view_tabs.count(), 1)
        self.assertEqual(w.view_tabs.currentIndex(), 0)
        self.assertFalse(w.close_all_tabs_btn.isEnabled())
        self.assertEqual(w.mode, "inventory")

    def test_tree_move_undo_redo_and_conflict(self):
        w = self.window
        target = w.store.add_node("group", "이동 위치")
        other = w.store.add_node("device", "추가 장치", self.group)
        w.refresh()
        w.move_nodes([self.device, self.units[0], other], target)
        self.assertEqual(w.store.node(self.units[0])["parent"], self.device)
        self.assertEqual(len(w.undo_moves), 1)
        w.tree.setFocus()
        APP.processEvents()
        QTest.keyClick(w.tree, Qt.Key.Key_Z, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(w.store.node(self.device)["parent"], self.group)
        self.assertEqual(w.store.node(other)["parent"], self.group)
        QTest.keyClick(w.tree, Qt.Key.Key_Y, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(w.store.node(self.device)["parent"], target)
        w.undo_tree_move()
        self.assertEqual(len(w.redo_moves), 1)
        w.move_nodes([self.device], None)
        self.assertEqual(len(w.redo_moves), 0)
        # A conflicting name must leave the undo command and all nodes unchanged.
        w.store.add_node("device", "다기능 통신장치", self.group)
        before = len(w.undo_moves)
        with self.assertRaises(Exception): w.undo_tree_move()
        self.assertIsNone(w.store.node(self.device)["parent"])
        self.assertEqual(len(w.undo_moves), before)
        w.close_project()
        self.assertEqual(w.undo_moves, [])
        self.assertEqual(w.view_tabs.count(), 1)

    def test_node_enter_saves_and_validates(self):
        w = self.window
        def submit(dialog):
            dialog.show()
            dialog.name.setFocus()
            dialog.name.setText("Enter 수정" if dialog.name.text() else "Enter 저장")
            APP.processEvents()
            QTest.keyClick(dialog.name, Qt.Key.Key_Return)
            self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
            return dialog.result()
        with patch.object(NodeDialog, "exec", submit):
            w.add_node("group")
        created = next(n for n in w.store.nodes() if n["name"] == "Enter 저장")
        w.select_node(created["id"])
        with patch.object(NodeDialog, "exec", submit):
            w.edit_node()
        self.assertEqual(w.store.node(created["id"])["name"], "Enter 수정")
        dialog = NodeDialog(w.store, "device", parent=w)
        dialog.show()
        dialog.name.setFocus()
        APP.processEvents()
        with patch.object(QMessageBox, "warning") as warning:
            QTest.keyClick(dialog.name, Qt.Key.Key_Return)
        warning.assert_called_once()
        self.assertTrue(dialog.isVisible())
        dialog.reject()

    def test_event_targets_saved_with_creation_and_edit(self):
        w = self.window
        w.select_node(self.group)
        def create(dialog):
            self.assertEqual(set(dialog.target_ids()), set(self.units))
            dialog.target_items[self.units[0]].setCheckState(0, Qt.CheckState.Unchecked)
            dialog.target_search.setText("COM-002")
            self.assertFalse(dialog.target_items[self.units[1]].isHidden())
            self.assertTrue(dialog.target_items[self.units[2]].isHidden())
            dialog.rich.editor.setPlainText("작성 중 대상 선택")
            dialog.show()
            dialog.location.setFocus()
            APP.processEvents()
            dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "event-picker.png"))
            QTest.keyClick(dialog.location, Qt.Key.Key_Return)
            self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
            return dialog.result()
        with patch.object(EntryDialog, "exec", create):
            w.new_event()
        ident = w.store.entries()[0]["id"]
        self.assertEqual(w.mode, "inventory")
        self.assertEqual(set(w.store.targets(ident)), set(self.units[1:]))
        extra = w.store.add_node("unit", "새 장비", self.device, serial="COM-007")
        def edit(dialog):
            self.assertNotIn(extra, dialog.target_ids())
            self.assertEqual(set(dialog.target_ids()), set(self.units[1:]))
            dialog.clear_targets()
            dialog.target_items[extra].setCheckState(0, Qt.CheckState.Checked)
            dialog.validate()
            return dialog.result()
        with patch.object(EntryDialog, "exec", edit):
            w.edit_entry(ident)
        self.assertEqual(w.store.targets(ident), [extra])
        before = len(w.store.entries(events_only=True))
        with patch.object(EntryDialog, "exec", lambda dialog: QDialog.DialogCode.Rejected):
            w.new_event()
        self.assertEqual(len(w.store.entries(events_only=True)), before)

    def test_welcome_and_compact_tree(self):
        w = self.window
        self.assertEqual(w.calendar_view_btn.text(), "달력 보기")
        self.assertIn("전체 기록 달력", w.calendar_view_btn.toolTip())
        self.assertEqual(w.project_pages.currentIndex(), 1)
        self.assertTrue(w.tree_items[self.group].isExpanded())
        self.assertFalse(w.tree_items[self.device].isExpanded())
        w.tree_items[self.device].setExpanded(True)
        height = w.tree.visualItemRect(w.tree_items[self.units[0]]).height()
        self.assertLessEqual(height, 26)
        w.close_project()
        self.assertEqual(w.project_pages.currentIndex(), 0)
        self.assertFalse(w.navigation.isVisible())
        APP.processEvents()
        w.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "welcome.png"))
        with patch("SEMP.QFileDialog.getOpenFileName", return_value=(str(self.project_file), "")):
            w.open_project()
        self.assertEqual(w.project_pages.currentIndex(), 1)
        self.assertEqual(len(w.store.nodes()), 8)

    def test_tree_sort_name_ascending_descending_and_created(self):
        w = self.window
        for name in ("B", "A10", "A2"):
            w.store.add_node("device", name, self.group)
        w.refresh()
        group_item = w.tree_items[self.group]
        child_names = lambda: [group_item.child(i).text(0) for i in range(group_item.childCount())]
        self.assertEqual(child_names(), ["A2", "A10", "B", "다기능 통신장치"])
        w.tree_sort.setCurrentIndex(w.tree_sort.findData("name_desc"))
        group_item = w.tree_items[self.group]
        self.assertEqual([group_item.child(i).text(0) for i in range(group_item.childCount())],
                         ["다기능 통신장치", "B", "A10", "A2"])
        w.tree_sort.setCurrentIndex(w.tree_sort.findData("created"))
        group_item = w.tree_items[self.group]
        self.assertEqual([group_item.child(i).text(0) for i in range(group_item.childCount())],
                         ["다기능 통신장치", "B", "A10", "A2"])

    def test_calendar_date_lists_records_then_create_button_opens_editor(self):
        w = self.window
        w.clear_selection()
        w.resize(1120, 720)
        APP.processEvents()
        self.assertTrue(w.calendar_panel.isVisible())
        self.assertFalse(w.metrics_panel.isVisible())
        self.assertFalse(w.tabs.isVisible())
        chosen = QDate(2026, 10, 15)
        w.calendar.setSelectedDate(chosen)
        def submit(dialog):
            self.assertFalse(hasattr(dialog, "title_edit"))
            self.assertEqual(dialog.day.date(), chosen)
            dialog.rich.editor.setPlainText("GCU 정기 점검 및 케이블 연결 확인")
            dialog.category.setCurrentText("정기 점검")
            dialog.category.setCurrentText("보정")
            dialog.category.setCurrentText("정기 점검")
            dialog.validate()
            return dialog.result()
        opened = []
        def submit_tracked(dialog):
            opened.append(True)
            return submit(dialog)
        with patch.object(EntryDialog, "exec", submit_tracked):
            view = w.calendar.findChild(QTableView)
            QTest.mouseClick(view.viewport(), Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, view.visualRect(view.currentIndex()).center())
            self.assertEqual(opened, [])
            self.assertIn("기록 없음", w.calendar_entries.item(0).text())
            QTest.mouseClick(w.calendar_create_btn, Qt.MouseButton.LeftButton)
        self.assertEqual(opened, [True])
        entry = w.store.entries()[0]
        self.assertEqual(entry["day"], "2026-10-15")
        self.assertEqual(entry["title"], "GCU 정기 점검 및 케이블 연결 확인")
        self.assertEqual(entry["category"], "정기 점검")
        self.assertIn("보정", w.store.categories())
        self.assertIn("정기 점검", [w.category_filter.itemText(i) for i in range(w.category_filter.count())])
        w.set_mode("inventory")
        w.clear_selection()
        APP.processEvents()
        self.assertEqual(len(w.calendar.events_by_day["2026-10-15"]), 1)
        self.assertLess(w.calendar.geometry().bottom(), w.calendar_day_label.geometry().top())
        w.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "calendar.png"))

    def test_tag_chips_date_picker_and_category_buttons(self):
        dialog = EntryDialog(self.window.store, event=True, parent=self.window)
        dialog.show()
        APP.processEvents()
        tag = dialog.tag_edit
        tag.input.setText("#현장A")
        QTest.keyClick(tag.input, Qt.Key.Key_Space)
        self.assertEqual(tag.tokens, ["#현장A"])
        self.assertEqual(tag.input.text(), "")
        tag.input.setText("#현장A #점검")
        QTest.keyClick(tag.input, Qt.Key.Key_Space)
        self.assertEqual(tag.tokens, ["#현장A", "#점검"])
        tag.input.setText("#미완성")
        self.assertEqual(tag.text(), "#현장A #점검 #미완성")
        QTest.mouseClick(tag.chips.findChildren(QPushButton)[0], Qt.MouseButton.LeftButton)
        self.assertNotIn("#현장A", tag.text())
        with patch.object(QInputDialog, "getText", return_value=("검교정", True)):
            QTest.mouseClick(dialog.category.add_button, Qt.MouseButton.LeftButton)
        self.assertEqual(dialog.category.currentText(), "검교정")
        QTest.mouseClick(dialog.calendar_button, Qt.MouseButton.LeftButton)
        self.assertTrue(dialog.mini_calendar.isVisible())
        self.assertEqual(dialog.calendar_button.text(), "")
        self.assertFalse(dialog.calendar_button.icon().isNull())
        dialog.mini_calendar.clicked.emit(QDate(2026, 11, 12))
        self.assertEqual(dialog.day.date(), QDate(2026, 11, 12))
        self.assertFalse(dialog.mini_calendar.isVisible())
        dialog.day.setFocus()
        dialog.day.setCurrentSection(dialog.day.Section.DaySection)
        QTest.keyClick(dialog.day, Qt.Key.Key_Up)
        self.assertEqual(dialog.day.date(), QDate(2026, 11, 13))
        dialog.rich.editor.setPlainText("검교정 결과 정상")
        APP.processEvents()
        dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "event-chips.png"))
        dialog.reject()
        self.assertNotIn("검교정", self.window.store.categories())

    def test_collection_editor_and_event_reuse(self):
        w = self.window
        other_group = w.store.add_node("group", "다른 그룹")
        other_device = w.store.add_node("device", "GCU", other_group)
        other_unit = w.store.add_node("unit", "GCU", other_device, serial="A001")
        w.refresh()
        self.assertEqual(node_caption(w.store.node(other_unit)), "GCU-A001")
        original_parent = w.store.node(other_unit)["parent"]
        def submit(dialog):
            dialog.name.setText("현장 시험 세트")
            for ident in (self.units[0], other_unit):
                dialog.picker.target_items[ident].setCheckState(0, Qt.CheckState.Checked)
            dialog.validate()
            return dialog.result()
        with patch.object(CollectionDialog, "exec", submit):
            w.new_collection()
        collection = w.store.collections()[0]
        self.assertEqual(set(w.store.collection_targets(collection["id"])), {self.units[0], other_unit})
        self.assertEqual(w.store.node(other_unit)["parent"], original_parent)
        dialog = EntryDialog(w.store, event=True, parent=w)
        picker = dialog.picker
        picker.collection_choice.setCurrentIndex(picker.collection_choice.findData(collection["id"]))
        picker.load_collection(w.store)
        picker.load_collection(w.store)
        self.assertEqual(set(dialog.target_ids()), {self.units[0], other_unit})
        dialog.rich.editor.setPlainText("모음 장비 점검")
        ident = w.store.save_entry(event=True, node_ids=dialog.target_ids(), **dialog.values())
        self.assertEqual(set(w.store.targets(ident)), {self.units[0], other_unit})
        dialog.close()
        w.store.delete_collection(collection["id"])
        self.assertEqual(set(w.store.targets(ident)), {self.units[0], other_unit})

    def test_recent_project_persists_and_missing_path_is_disabled(self):
        w = self.window
        self.assertEqual(w.recent_paths(), [str(self.project_file)])
        missing = str(Path(self.tmp.name) / "missing" / "project.semp")
        w.settings.setValue("recentProjects", [str(self.project_file), missing])
        w.refresh_recent()
        self.assertFalse(w.recent_list.item(1).flags() & Qt.ItemFlag.ItemIsEnabled)
        w.close_project()
        w.refresh_recent()
        APP.processEvents()
        QTest.mouseClick(w.recent_list.viewport(), Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, w.recent_list.visualItemRect(w.recent_list.item(0)).center())
        self.assertEqual(w.store.path, self.project_file)
        self.assertEqual(w.recent_paths().count(str(self.project_file)), 1)
        other = MainWindow(settings=QSettings(str(Path(self.tmp.name) / "settings.ini"), QSettings.Format.IniFormat))
        self.assertEqual(other.recent_list.item(0).data(Qt.ItemDataRole.UserRole), str(self.project_file))
        other.close()

    def test_recent_location_and_tag_suggestions_fill_fields(self):
        w = self.window
        for index in range(7):
            w.store.save_entry(event=True, title=f"기록 {index}", plain="점검 결과", html="<p>점검 결과</p>", day="2026-09-09", category="기타", location=f"현장 {index}", tag_text=f"#위치{index}")
        dialog = EntryDialog(w.store, event=True, parent=w)
        dialog.show()
        APP.processEvents()
        self.assertEqual(dialog.windowTitle(), "기록 작성")
        self.assertEqual(dialog.location.placeholderText(), "현 위치(선택)")
        self.assertEqual(len(dialog.location_suggestions.buttons), 5)
        self.assertEqual(len(dialog.tag_edit.suggestions.buttons), 5)
        QTest.mouseClick(dialog.location_suggestions.buttons[0], Qt.MouseButton.LeftButton)
        self.assertEqual(dialog.location.text(), "현장 6")
        dialog.tag_edit.input.setText("#작성중")
        QTest.mouseClick(dialog.tag_edit.suggestions.buttons[0], Qt.MouseButton.LeftButton)
        self.assertEqual(dialog.tag_edit.tokens, ["#작성중", "#위치6"])
        QTest.mouseClick(dialog.tag_edit.suggestions.buttons[0], Qt.MouseButton.LeftButton)
        self.assertEqual(len(dialog.tag_edit.tokens), 2)
        self.assertFalse(dialog.picker.collection_apply.isEnabled())
        dialog.rich.editor.setPlainText("현장 장비 점검 완료")
        APP.processEvents()
        dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "record-recommendations.png"))
        dialog.reject()

    def test_palette_grid_applies_to_selection_and_cancel_keeps_format(self):
        dialog = EntryDialog(self.window.store, event=True, parent=self.window)
        editor = dialog.rich.editor
        editor.setPlainText("선택 일반")
        cursor = editor.textCursor()
        cursor.setPosition(0)
        cursor.setPosition(2, QTextCursor.MoveMode.KeepAnchor)
        editor.setTextCursor(cursor)
        def select_color(palette):
            palette.show()
            APP.processEvents()
            self.assertEqual(len(palette.swatches), 16)
            palette.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "text-palette.png"))
            QTest.mouseClick(palette.swatches[13], Qt.MouseButton.LeftButton)
            return palette.result()
        with patch.object(ColorPaletteDialog, "exec", select_color):
            dialog.rich.choose_color("color")
        self.assertEqual(editor.textCursor().charFormat().foreground().color().name(), "#1d4ed8")
        before = editor.toHtml()
        with patch.object(ColorPaletteDialog, "exec", lambda palette: QDialog.DialogCode.Rejected):
            dialog.rich.choose_color("color")
        self.assertEqual(editor.toHtml(), before)
        palette = ColorPaletteDialog("highlight", QColor("#ffffff"), dialog)
        self.assertEqual(len(palette.swatches), 24)
        palette.hex_edit.setText("invalid")
        palette.apply_hex()
        self.assertEqual(palette.result(), QDialog.DialogCode.Rejected)
        self.assertTrue(palette.error.text())
        palette.hex_edit.setText("#FDE68A")
        palette.apply_hex()
        self.assertEqual(palette.selected_color.name(), "#fde68a")
        dialog.rich.apply_format("highlight", palette.selected_color)
        cursor.setPosition(4)
        editor.setTextCursor(cursor)
        self.assertNotEqual(cursor.charFormat().foreground().color().name(), "#1d4ed8")
        self.assertNotEqual(cursor.charFormat().background().color().name(), "#fde68a")
        palette.close()
        dialog.close()

    def test_device_description_and_serial_purpose_are_separate(self):
        w = self.window
        device = NodeDialog(w.store, "device", parent=w)
        device.show()
        APP.processEvents()
        self.assertFalse(device.purpose.isVisible())
        self.assertTrue(device.description.isVisible())
        device.name.setText("GCU")
        device.description.setPlainText("비행 제어 장치 설명")
        values = device.values()
        self.assertEqual(values["purpose"], "")
        self.assertEqual(values["description"], "비행 제어 장치 설명")
        ident = w.store.add_node("device", parent=self.group, **values)
        device.close()
        w.refresh()
        w.select_node(ident)
        def register(dialog):
            first = dialog.findChild(QLineEdit, "SerialFirst")
            purpose = dialog.findChild(QComboBox, "UnitPurpose")
            description = dialog.findChild(QTextEdit, "UnitDescription")
            self.assertEqual(description.toPlainText(), "비행 제어 장치 설명")
            first.setText("A002")
            purpose.addItem("예비 시험용", "예비 시험용")
            purpose.setCurrentIndex(purpose.findData("예비 시험용"))
            dialog.rebuild()
            dialog.apply_all()
            description.setPlainText("개별 장비 점검 설명")
            save = next(b for b in dialog.findChildren(QPushButton) if b.text() == "등록")
            save.click()
            return dialog.result()
        with patch.object(QDialog, "exec", register):
            w.add_serials()
        unit = next(n for n in w.store.nodes() if n["parent"] == ident)
        self.assertEqual(unit["purpose"], "예비 시험용")
        self.assertEqual(unit["description"], "개별 장비 점검 설명")
        self.assertEqual(w.store.node(ident)["description"], "비행 제어 장치 설명")
        edit = NodeDialog(w.store, "unit", unit, w)
        edit.show()
        APP.processEvents()
        self.assertTrue(edit.purpose.isVisible())
        self.assertTrue(edit.description.isVisible())
        edit.close()
        w.select_node(unit["id"])
        self.assertEqual(w.device_badge.text(), "GCU")
        self.assertEqual(w.detail_title.text(), "A002")
        self.assertEqual(w.tree_items[unit["id"]].text(0), "A002")
        w.resize(1120, 720)
        APP.processEvents()
        w.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "device-badge.png"))

    def test_purpose_change_from_record_and_detail(self):
        w = self.window
        second_device = w.store.add_node("device", "보조 장치", self.group)
        second_unit = w.store.add_node("unit", "보조 장치", second_device, serial="SUB-001", purpose="예비용")
        w.refresh()
        w.select_node(self.device)
        w.tree_items[second_device].setSelected(True)
        def write_change(dialog):
            dialog.show()
            APP.processEvents()
            dialog.category.choices["용도변경"].click()
            self.assertTrue(dialog.purpose_change_panel.isVisible())
            dialog.purpose_after.setCurrentText("환경용")
            dialog.rich.editor.setPlainText("운용 구분 변경")
            dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "purpose-change-record.png"))
            dialog.validate()
            return dialog.result()
        with patch.object(EntryDialog, "exec", write_change):
            w.new_event()
        self.assertEqual({w.store.node(ident)["purpose"] for ident in [*self.units, second_unit]}, {"환경용"})
        saved = w.store.entries(events_only=True)[0]
        self.assertEqual(saved["category"], "용도변경")
        self.assertEqual(saved["purpose_after"], "환경용")

        w.select_node(second_unit)
        self.assertIn("더블클릭", w.metrics[2][1].toolTip())
        def change_detail(dialog):
            self.assertEqual(dialog.windowTitle(), "용도 변경")
            dialog.purpose.setCurrentText("점검용")
            dialog.validate()
            return dialog.result()
        with patch.object(PurposeChangeDialog, "exec", change_detail):
            w.metrics[2][1].doubleClicked.emit()
        self.assertEqual(w.store.node(second_unit)["purpose"], "점검용")

    def test_record_purpose_choices_scroll_when_vocabulary_is_long(self):
        w = self.window
        w.store.db.executemany("INSERT OR IGNORE INTO purposes VALUES(?)",
                               [(f"추가 용도 {index:02}",) for index in range(35)])
        dialog = EntryDialog(w.store, event=True, parent=w)
        dialog.show()
        dialog.category.choices["용도변경"].click()
        APP.processEvents()
        self.assertTrue(dialog.purpose_scroll.isVisible())
        self.assertLessEqual(dialog.purpose_scroll.height(), 150)
        self.assertGreater(dialog.purpose_scroll.verticalScrollBar().maximum(), 0)
        dialog.close()

    def test_custom_purpose_candidate_can_be_removed_without_changing_equipment(self):
        w = self.window
        unit = w.store.node(self.units[0])
        w.store.update_node(unit["id"], unit["stored_name"], unit["serial"], "오타용",
                            unit["description"], unit["tags"], purpose_options=["오타용"])
        dialog = EntryDialog(w.store, event=True, parent=w)
        self.assertIn("오타용", dialog.purpose_after.choices)
        self.assertTrue(dialog.purpose_after.remove_choice("오타용"))
        self.assertNotIn("오타용", w.store.vocab("purpose"))
        self.assertEqual(w.store.node(unit["id"])["purpose"], "오타용")
        self.assertFalse(dialog.purpose_after.remove_choice("비행용"))
        dialog.close()

    def test_record_preselection_only_from_inventory_detail(self):
        w = self.window
        w.select_node(self.units[0])
        cases = [("inventory", None, {self.units[0]}), ("search", None, set()), ("events", None, set()), ("inventory", QDate(2026, 10, 1), set())]
        for mode, day, expected in cases:
            w.set_mode(mode)
            def inspect(dialog):
                self.assertEqual(set(dialog.target_ids()), expected)
                return QDialog.DialogCode.Rejected
            with patch.object(EntryDialog, "exec", inspect):
                w.new_event(day=day)
        self.assertNotIn("events", w.nav_buttons)
        w.set_mode("inventory")
        w.clear_selection()
        def inspect_empty(dialog):
            self.assertEqual(dialog.target_ids(), [])
            return QDialog.DialogCode.Rejected
        with patch.object(EntryDialog, "exec", inspect_empty):
            w.new_event()

    def test_tree_duplicate_shortcut_copies_subtree(self):
        w = self.window
        w.select_node(self.device)
        QTest.keyClick(w.tree, Qt.Key.Key_D, Qt.KeyboardModifier.ControlModifier)
        copied = w.store.node(w.selected_id)
        self.assertEqual(copied["name"], "다기능 통신장치 (1)")
        children = [n for n in w.store.nodes() if n["parent"] == copied["id"]]
        self.assertEqual(len(children), 6)
        w.select_node(children[0]["id"])
        self.assertEqual(w.device_badge.text(), copied["name"])
        self.assertEqual(w.detail_title.text(), children[0]["serial"])
        self.assertEqual(w.tree_items[children[0]["id"]].text(0), children[0]["serial"])

    def test_purpose_buttons_defaults_custom_choices_and_reopen(self):
        w = self.window
        unit = w.store.node(self.units[0])
        dialog = NodeDialog(w.store, "unit", unit, w)
        dialog.show()
        APP.processEvents()
        purpose = dialog.purpose
        self.assertEqual(list(purpose.choices)[:3], ["비행용", "예비용", "환경용"])
        self.assertEqual(purpose.currentText(), unit["purpose"])
        for custom in ("지상시험용", "교육용"):
            with patch.object(QInputDialog, "getText", return_value=(custom, True)):
                QTest.mouseClick(purpose.add_button, Qt.MouseButton.LeftButton)
            self.assertEqual(purpose.currentText(), custom)
        QTest.mouseClick(purpose.choices["비행용"], Qt.MouseButton.LeftButton)
        self.assertEqual(purpose.currentText(), "비행용")
        self.assertEqual(sum(b.isChecked() for b in purpose.choices.values()), 1)
        self.assertIs(purpose.flow.itemAt(purpose.flow.count() - 1).widget(), purpose.add_button)
        w.store.update_node(unit["id"], **dialog.values())
        self.assertEqual(w.store.node(unit["id"])["purpose"], "비행용")
        dialog.close()
        reopened = NodeDialog(w.store, "unit", w.store.node(unit["id"]), w)
        self.assertTrue(reopened.purpose.choices["비행용"].isChecked())
        self.assertIn("지상시험용", reopened.purpose.choices)
        self.assertIn("교육용", reopened.purpose.choices)
        reopened.show()
        APP.processEvents()
        reopened.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "purpose-buttons.png"))
        reopened.purpose.setCurrentText("취소한 용도")
        reopened.reject()
        self.assertNotIn("취소한 용도", w.store.vocab("purpose"))
        blank = PurposeButtons()
        self.assertEqual(blank.currentText(), "")
        blank.close()

    def test_logo_and_project_name_return_to_calendar_home(self):
        w = self.window
        w.resize(1120, 720)
        APP.processEvents()
        self.assertLess(w.logo_link.geometry().right(), w.nav_buttons["inventory"].geometry().left())
        for link, mode in ((w.logo_link, "search"), (w.project_label, "inventory"), (w.project_label, "events")):
            w.select_node(self.units[0])
            w.set_mode(mode)
            QTest.mouseClick(link, Qt.MouseButton.LeftButton)
            APP.processEvents()
            self.assertEqual(w.mode, "inventory")
            self.assertIsNone(w.selected_id)
            self.assertEqual(w.tree.ids(), [])
            self.assertTrue(w.calendar_panel.isVisible())
            self.assertFalse(w.tabs.isVisible())
        w.select_node(self.device)
        w.logo_link.setFocus()
        QTest.keyClick(w.logo_link, Qt.Key.Key_Return)
        self.assertIsNone(w.selected_id)
        self.assertTrue(w.calendar_panel.isVisible())

    def test_calendar_counts_legacy_and_shared_records_once(self):
        w = self.window
        for event, targets in ((False, [self.units[0]]), (True, self.units[:2])):
            w.store.save_entry(title="본문 반복 없음", plain="본문 반복 없음", html="<p>본문 반복 없음</p>", category="위치", day="2026-09-09", event=event, node_ids=targets)
        w.refresh()
        w.show_home()
        w.calendar.setSelectedDate(QDate(2026, 9, 9))
        self.assertEqual(len(w.calendar.events_by_day["2026-09-09"]), 2)
        self.assertEqual(w.calendar_entries.count(), 2)
        w.select_node(self.units[0])
        card = w.timeline.itemWidget(w.timeline.item(0))
        texts = [label.text() for label in card.findChildren(QLabel)]
        self.assertEqual(texts.count("본문 반복 없음"), 1)
        self.assertIn("위치", texts)
        w.show_home()
        APP.processEvents()
        w.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "calendar-all-records.png"))

    def test_purpose_ordinal_filters_and_filtered_group_check(self):
        w = self.window
        d = w.store.add_node("device", "추가 장치", self.group)
        u2 = w.store.add_node("unit", "추가 장치", d, serial="A2", purpose="점검용")
        u10 = w.store.add_node("unit", "추가 장치", d, serial="A10", purpose="비행용")
        w.store.update_node(self.units[0], "다기능 통신장치", "COM-001", "점검용", "", "")
        w.refresh()
        w.equipment_filters.purpose.setCurrentIndex(w.equipment_filters.purpose.findData("점검용"))
        w.equipment_filters.ordinal.setValue(1)
        shown = {ident for ident in [*self.units, u2, u10] if not w.tree_items[ident].isHidden()}
        self.assertEqual(shown, {self.units[0]})
        dialog = CollectionDialog(w.store, parent=w)
        picker = dialog.picker
        picker.filters.ordinal_buttons[1].click()
        picker.filters.purpose_buttons["점검용"].click()
        picker.target_items[self.group].setCheckState(0, Qt.CheckState.Checked)
        self.assertEqual(set(picker.target_ids()), {self.units[0]})
        picker.clear_targets()
        picker.filters.reset()
        picker.filters.ordinal_buttons[2].click()
        picker.select_visible_targets()
        self.assertEqual(set(picker.target_ids()), {self.units[1], u2})
        picker.clear_targets()
        picker.filters.reset()
        picker.filters.purpose_buttons["점검용"].click()
        picker.filters.purpose_buttons["비행용"].click()
        picker.filters.ordinal_buttons[1].click()
        picker.filters.kind_buttons["unit"].click()
        self.assertEqual(set(picker.filters.selected(picker.filters.purpose_buttons)), {"점검용", "비행용"})
        self.assertEqual(picker.filters.matches(w.store.nodes()), {self.units[0]})
        self.assertIn("점검용 2", picker.filters.purpose_buttons["점검용"].text())
        picker.filters.reset()
        picker.filters.kind_buttons["device"].click()
        self.assertEqual(picker.filters.matches(w.store.nodes(), "추가 장치"), {u2, u10})
        picker.filters.reset()
        picker.filters.ordinal_input.setValue(7)
        picker.filters.add_ordinal()
        self.assertTrue(picker.filters.ordinal_buttons[7].isChecked())
        self.assertIn("7번째", picker.filters.ordinal_buttons[7].text())
        year_device = w.store.add_node("device", "연도별 장치", self.group)
        year_units = [w.store.add_node("unit", "연도별 장치", year_device, serial=value)
                      for value in ("2024-GCU-A001", "2024-GCU-A002", "GCU-A001-2025", "GCU-A002-2025",
                                    "T-26001", "T-25001")]
        w.refresh()
        picker.filters.refresh(w.store.nodes())
        picker.filters.reset()
        picker.target_search.setText("연도별 장치")
        picker.filters.ordinal_buttons[1].click()
        self.assertEqual(picker.filters.matches(w.store.nodes(), "연도별 장치"),
                         {year_units[0], year_units[2], year_units[4], year_units[5]})
        w.equipment_filters.reset()
        w.tree_filter.setText("연도별 장치")
        w.equipment_filters.ordinal.setValue(1)
        self.assertEqual({ident for ident in year_units if not w.tree_items[ident].isHidden()},
                         {year_units[0], year_units[2], year_units[4], year_units[5]})
        dialog.name.setText("각 장치의 두 번째 장비")
        dialog.show()
        APP.processEvents()
        dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "collection-filters.png"))
        dialog.reject()

    def test_import_embedded_tree_create_group_and_indent_drop(self):
        w = self.window
        nested = w.store.add_node("group", "하위 등록 위치", self.group)
        w.refresh()
        dialog = ImportDialog(w)
        w.import_dialog = dialog
        dialog.show()
        dialog.rows = [ImportRow(2, "GCU", "001", "002", ["001", "002"], purpose="점검용")]
        dialog.render()
        dialog.table.selectRow(0)
        APP.processEvents()
        target = dialog.import_items[nested]
        point = dialog.import_tree.visualItemRect(target).center()
        parent_rect = dialog.import_tree.visualItemRect(dialog.import_items[self.group])
        self.assertIs(dialog.import_tree.resolve_drop_target(QPoint(parent_rect.left() + 1, point.y()), True), dialog.import_items[self.group])
        mime = QMimeData()
        mime.setData(IMPORT_MIME, json.dumps({"session": dialog.session, "rows": [0]}).encode())
        enter = QDragEnterEvent(point, Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        APP.sendEvent(dialog.import_tree.viewport(), enter)
        move = QDragMoveEvent(point, Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        APP.sendEvent(dialog.import_tree.viewport(), move)
        self.assertIs(dialog.import_tree.drop_target, target)
        self.assertIn("하위 등록 위치", dialog.drop_hint.text())
        APP.processEvents()
        dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "import-tree-drop.png"))
        drop = QDropEvent(QPointF(point), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        APP.sendEvent(dialog.import_tree.viewport(), drop)
        self.assertIn(0, dialog.imported)
        device = next(n for n in w.store.nodes() if n["name"] == "GCU" and n["kind"] == "device")
        self.assertEqual(device["parent"], nested)
        units = [n for n in w.store.nodes() if n["parent"] == device["id"]]
        self.assertEqual([n["purpose"] for n in units], ["점검용", "점검용"])
        dialog.import_tree.setCurrentItem(dialog.import_items[nested])
        def create_group(editor):
            editor.name.setText("새 폴더")
            editor.accept()
            return editor.result()
        with patch.object(NodeDialog, "exec", create_group): dialog.add_import_group()
        created = next(n for n in w.store.nodes() if n["name"] == "새 폴더")
        self.assertEqual(created["parent"], nested)
        self.assertIn(created["id"], dialog.import_items)
        dialog.import_tree.setCurrentItem(dialog.import_items[nested])
        def create_device(editor):
            editor.name.setText("빈 장치")
            editor.accept()
            return editor.result()
        with patch.object(NodeDialog, "exec", create_device): dialog.add_import_node("device")
        empty_device = next(n for n in w.store.nodes() if n["name"] == "빈 장치")
        self.assertEqual(empty_device["kind"], "device")
        self.assertEqual(empty_device["parent"], nested)
        self.assertIn(empty_device["id"], dialog.import_items)
        dialog.close()

    def test_target_picker_prioritizes_tall_scrollable_tree(self):
        dialog = CollectionDialog(self.window.store, parent=self.window)
        dialog.show()
        APP.processEvents()
        tree = dialog.picker.target_tree
        self.assertGreaterEqual(tree.height(), 500)
        self.assertLess(tree.geometry().left(), dialog.picker.filters.geometry().left())
        self.assertEqual(tree.verticalScrollBarPolicy(), Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.assertTrue(dialog.picker.target_items[self.group].isExpanded())
        self.assertFalse(dialog.picker.target_items[self.device].isExpanded())
        dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "collection-tall-tree.png"))
        dialog.reject()

    def test_selection_search_and_rich_format(self):
        w = self.window
        w.tree_items[self.device].setExpanded(True)
        w.select_node(self.units[0])
        self.assertEqual(w.detail_title.text(), "COM-001")
        self.assertEqual(w.device_badge.text(), "다기능 통신장치")
        self.assertTrue(w.device_badge.isVisible())
        self.assertEqual(w.tree_items[self.units[0]].text(0), "COM-001")
        p1 = w.tree.visualItemRect(w.tree_items[self.units[0]]).center()
        p2 = w.tree.visualItemRect(w.tree_items[self.units[1]]).center()
        p4 = w.tree.visualItemRect(w.tree_items[self.units[3]]).center()
        QTest.mouseClick(w.tree.viewport(), Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, p1)
        QTest.mouseClick(w.tree.viewport(), Qt.MouseButton.LeftButton, Qt.KeyboardModifier.ControlModifier, p2)
        self.assertEqual(len(w.tree.ids()), 2)
        QTest.mouseClick(w.tree.viewport(), Qt.MouseButton.LeftButton, Qt.KeyboardModifier.ShiftModifier, p4)
        self.assertGreaterEqual(len(w.tree.ids()), 3)
        dialog = EntryDialog(w.store, parent=w)
        dialog.rich.editor.setPlainText("선택 서식 시험")
        editor = dialog.rich.editor
        editor.setPlainText("중요 일반")
        cursor = editor.textCursor()
        cursor.setPosition(0)
        cursor.setPosition(2, QTextCursor.MoveMode.KeepAnchor)
        editor.setTextCursor(cursor)
        dialog.rich.apply_format("bold", True)
        ident = w.store.save_entry(node_ids=[self.units[0]], **dialog.values())
        self.assertIn("font-weight:700", w.store.entry(ident)["html"])
        cursor.setPosition(4)
        self.assertLess(cursor.charFormat().fontWeight(), 700)
        w.search.setText("#현장A")
        w.set_mode("search")
        self.assertGreaterEqual(w.search_results.topLevelItemCount(), 8)
        dialog.close()

    def test_event_drop_and_delete_unlinks_only(self):
        w = self.window
        e = w.store.save_entry(title="현장 시험", category="사용", day="2026-09-01", html="<b>통신 시험</b>", plain="통신 시험", event=True)
        w.event_id = e
        w.refresh()
        w.set_mode("events")
        mime = QMimeData()
        mime.setData(NODE_MIME, json.dumps([self.device, self.units[0]]).encode())
        enter = QDragEnterEvent(QPoint(20, 20), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        APP.sendEvent(w.targets.viewport(), enter)
        drop = QDropEvent(QPointF(20,20), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        APP.sendEvent(w.targets.viewport(), drop)
        self.assertEqual(w.targets.count(), 6)
        w.targets.setCurrentRow(0)
        removed = w.targets.currentItem().data(Qt.ItemDataRole.UserRole)
        QTest.keyClick(w.targets, Qt.Key.Key_Delete)
        self.assertEqual(w.targets.count(), 5)
        self.assertEqual(w.store.node(removed)["kind"], "unit")

    def test_import_click_and_tree_drop(self):
        w = self.window
        dialog = ImportDialog(w)
        w.import_dialog = dialog
        dialog.rows = [ImportRow(2,"케이블", "001", "003", ["001", "002", "003"]), ImportRow(3,"전원장치", "", "", [])]
        dialog.render()
        dialog.table.selectRow(0)
        dialog.destination.setCurrentIndex(dialog.destination.findData(self.group))
        dialog.register_selected()
        self.assertIn(0, dialog.imported)
        mime = QMimeData()
        mime.setData(IMPORT_MIME, json.dumps({"session": dialog.session, "rows": [1]}).encode())
        point = w.tree.visualItemRect(w.tree_items[self.group]).center()
        enter = QDragEnterEvent(point, Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        APP.sendEvent(w.tree.viewport(), enter)
        drop = QDropEvent(QPointF(point), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        APP.sendEvent(w.tree.viewport(), drop)
        self.assertIn(1, dialog.imported)
        self.assertEqual(sum(n["kind"]=="unit" for n in w.store.nodes()),9)

    def test_import_cells_can_be_edited_copied_pasted_and_registered(self):
        w = self.window
        dialog = ImportDialog(w)
        dialog.rows = [
            ImportRow(2, "원본A", "001", "001", ["001"]),
            ImportRow(3, "원본B", "002", "002", ["002"]),
        ]
        dialog.render()
        self.assertEqual(dialog.table.verticalHeader().defaultSectionSize(), 32)
        dialog.show()
        dialog.table.editItem(dialog.table.item(0, 0))
        APP.processEvents()
        editor = dialog.table.findChild(QLineEdit, "ImportCellEditor")
        self.assertIsNotNone(editor)
        self.assertLessEqual(editor.height(), 26)
        dialog.grab().save(str(Path(__file__).resolve().parent.parent / ".qa" / "import-cell-editor.png"))
        QApplication.clipboard().setText(
            "수정장치\t010\t012\t비행용\t격납고 A\t수정 설명\t#주력\n"
            "수정장치\t013\t\t예비용\t현장 B\t두 번째 설명\t예비")
        dialog.table.setCurrentCell(0, 0)
        QTest.keyClick(dialog.table, Qt.Key.Key_V, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(dialog.rows[0].serials, ["010", "011", "012"])
        self.assertEqual(dialog.rows[0].description, "수정 설명")
        self.assertEqual(dialog.rows[1].serials, ["013"])
        self.assertEqual(dialog.rows[1].tag_text, "#예비")
        dialog.table.clearSelection()
        for column in range(2):
            dialog.table.item(0, column).setSelected(True)
        QTest.keyClick(dialog.table, Qt.Key.Key_C, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(QApplication.clipboard().text(), "수정장치\t010")
        dialog.register([0, 1], self.group)
        device = next(node for node in w.store.nodes() if node["name"] == "수정장치")
        units = [node for node in w.store.nodes() if node["parent"] == device["id"]]
        self.assertEqual([node["serial"] for node in units], ["010", "011", "012", "013"])
        self.assertEqual(units[0]["location"], "격납고 A")
        self.assertEqual(units[-1]["description"], "두 번째 설명")
        dialog.close()

    def test_import_destination_tree_keeps_scroll_and_expansion_after_refresh(self):
        w = self.window
        groups = [w.store.add_node("group", f"등록 위치 {index:02}") for index in range(40)]
        dialog = ImportDialog(w)
        w.import_dialog = dialog
        dialog.show()
        APP.processEvents()
        dialog.import_items[self.group].setExpanded(False)
        scrollbar = dialog.import_tree.verticalScrollBar()
        self.assertGreater(scrollbar.maximum(), 0)
        scrollbar.setValue(scrollbar.maximum())
        previous = scrollbar.value()
        dialog.refresh_destinations()
        APP.processEvents()
        self.assertFalse(dialog.import_items[self.group].isExpanded())
        self.assertGreaterEqual(scrollbar.value(), previous)
        dialog.close()
        w.import_dialog = None

    def test_template_save_does_not_require_bundled_file(self):
        destination = Path(self.tmp.name) / "standalone-template.xlsx"
        with patch("SEMP.QFileDialog.getSaveFileName", return_value=(str(destination), "Excel (*.xlsx)")):
            self.window.save_template()
        self.assertTrue(destination.is_file())
        self.assertEqual(read_import(destination), [])

    def test_delete_key_trash_and_restore(self):
        w = self.window
        w.select_node(self.device)
        with patch.object(QMessageBox, "question", return_value=QMessageBox.StandardButton.Yes):
            QTest.keyClick(w.tree, Qt.Key.Key_Delete)
        self.assertEqual(len(w.store.nodes()),1)
        w.store.restore(w.store.trash()[0]["id"])
        w.refresh()
        self.assertEqual(len(w.tree_items),8)

    def test_render_desktop_views(self):
        w = self.window
        out = Path(__file__).resolve().parent.parent / ".qa"
        out.mkdir(exist_ok=True)
        examples = [("2026-09-08", "입고", "현장 시험 종료 · 장비 입고", "외관과 연결부를 확인했습니다. 점검 후 보관함에 입고했습니다.", "장비 보관실"), ("2026-09-06", "이슈", "통신 연결부 접촉 상태 확인", "커넥터 고정 상태 재확인 필요. 다음 시험 전 점검 예정입니다.", ""), ("2026-09-04", "출고", "현장 A로 장비 반출", "시험 지원을 위해 장비를 반출했습니다. 케이블 세트를 함께 준비했습니다.", "현장 A")]
        for day, category, title, plain, location in examples:
            w.store.save_entry(title=title,category=category,day=day,html=f"<p>{plain}</p>",plain=plain,tag_text="#현장A #시험지원",location=location,node_ids=[self.units[0]])
        w.store.save_entry(title="통신장치 공통 점검",category="기타",day="2026-09-02",html="<p>전체 장비 펌웨어 버전 확인 완료.</p>",plain="전체 장비 펌웨어 버전 확인 완료.",node_ids=[self.device])
        event = w.store.save_entry(title="9월 현장 통합 시험",category="사용",day="2026-09-04",html="<h3>현장 A · 통신 연동 시험</h3><p>장비 간 연결과 데이터 수신 상태를 확인합니다.</p><p><b>준비 사항</b></p><p>전원 공급 상태 확인<br>통신 케이블 및 예비 장비 준비</p>",plain="현장 통신 연동 시험",tag_text="#현장A #통합시험",event=True)
        w.store.link_event(event, [self.group])
        w.refresh()
        w.select_node(self.units[0])
        APP.processEvents()
        w.grab().save(str(out / "inventory.png"))
        w.event_id = event
        w.refresh_events()
        w.set_mode("events")
        APP.processEvents()
        w.grab().save(str(out / "events.png"))
        w.set_mode("inventory")
        w.resize(1120,720)
        APP.processEvents()
        w.grab().save(str(out / "compact.png"))
        dialog = EntryDialog(w.store,parent=w)
        dialog.rich.editor.setPlainText("현장 시험 종료 · 장비 입고")
        dialog.rich.editor.setHtml("<p>장비 상태를 확인했습니다.</p><p><b>확인 사항</b></p><p>외관 이상 없음<br><span style='color:#15816f'>통신 연결 상태 정상</span></p>")
        dialog.show()
        APP.processEvents()
        dialog.grab().save(str(out / "editor.png"))
        dialog.close()


if __name__ == "__main__": unittest.main()
