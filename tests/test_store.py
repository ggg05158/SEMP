import shutil
import sqlite3
import tempfile
import unittest
from contextlib import closing
from unittest.mock import patch
from pathlib import Path
from SEMP import Store, serial_range, tags


class StoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.store = Store.create(self.tmp.name, "시험 프로젝트", default_groups=False)
        self.group = self.store.add_node("group", "장비")
        self.device = self.store.add_node("device", "통신장치", self.group)
        self.unit = self.store.add_node("unit", "통신장치", self.device, "001")

    def tearDown(self):
        self.store.close()
        self.tmp.cleanup()

    def record(self, target=None, **kwargs):
        values = dict(title="시험", category="기타", day="2026-01-01", html="<b>메모</b>", plain="메모", node_ids=[target or self.unit])
        values.update(kwargs)
        return self.store.save_entry(**values)

    def test_default_groups_and_existing_project_upgrade(self):
        project = Store.create(self.tmp.name, "새 프로젝트")
        self.assertEqual([n["name"] for n in project.nodes()], ["탑재장치", "점검장비", "케이블", "기타"])
        path = project.path
        project.db.executescript("DROP TABLE collection_nodes; DROP TABLE collections; DROP TABLE categories;")
        project.close()
        project = Store(path)
        self.assertEqual(len(project.nodes()), 4)
        self.assertEqual(project.collections(), [])
        self.assertIn("기타", project.categories())
        project.close()

    def test_import_saves_extended_unit_fields(self):
        created, count = self.store.import_rows(self.group, [
            ("GCU", ["2501", "2502"], "실전용", "격납고 A", "비행 제어용", "#주력 #비행")
        ])
        self.assertEqual(count, 2)
        device = next(node for node in self.store.nodes() if node["name"] == "GCU" and node["kind"] == "device")
        self.assertEqual(device["parent"], self.group)
        units = [node for node in self.store.nodes() if node["parent"] == device["id"]]
        self.assertEqual([node["serial"] for node in units], ["2501", "2502"])
        self.assertTrue(all(node["purpose"] == "실전용" for node in units))
        self.assertTrue(all(node["location"] == "격납고 A" for node in units))
        self.assertTrue(all(node["description"] == "비행 제어용" for node in units))
        self.assertTrue(all(node["tags"] == "#주력 #비행" for node in units))
        self.assertEqual(self.store.current_state(units[0]["id"])[0], "격납고 A")

    def test_collection_portability_custom_categories_and_deleted_members(self):
        s = self.store
        collection = s.save_collection("시험 세트", [self.unit])
        self.record(category="검교정", category_options=["검교정", "보정"])
        s.delete([self.unit])
        self.assertEqual(s.collection_targets(collection), [])
        s.restore(s.trash()[0]["id"])
        self.assertEqual(s.collection_targets(collection), [self.unit])
        s.close()
        target = Path(self.tmp.name) / "옮긴 프로젝트"
        shutil.copytree(s.root, target)
        self.store = Store(target / "project.semp")
        self.assertEqual(self.store.collection_targets(collection), [self.unit])
        self.assertEqual(self.store.categories()[-2:], ["검교정", "보정"])
        self.assertEqual(self.store.node(self.unit)["parent"], self.device)

    def test_recent_inputs_order_limit_edit_and_reopen(self):
        ids = []
        for index in range(7):
            ids.append(self.record(location=f"장소 {index}", tag_text=f"#태그{index}"))
        self.assertEqual(self.store.recent_suggestions("location"), [f"장소 {i}" for i in range(6, 1, -1)])
        entry = self.store.entry(ids[0])
        self.store.save_entry(ident=ids[0], title=entry["title"], category=entry["category"], day=entry["day"], html=entry["html"], plain=entry["plain"], location="장소 0", tag_text="#태그0", node_ids=[self.unit])
        self.assertEqual(self.store.recent_suggestions("location")[0], "장소 0")
        self.assertEqual(self.store.recent_suggestions("tags")[0], "#태그0")
        before = self.store.recent_suggestions("location")
        with self.assertRaises(ValueError):
            self.store.save_entry(title="", category="기타", day="2026-09-09", html="", plain="", location="실패한 장소")
        self.assertEqual(self.store.recent_suggestions("location"), before)
        path = self.store.path
        self.store.close()
        self.store = Store(path)
        self.assertEqual(self.store.recent_suggestions("location"), before)
        self.store.db.execute("DELETE FROM recent_inputs")
        self.assertEqual(self.store.recent_suggestions("location")[0], "장소 6")

    def test_duplicate_subtrees_names_and_independent_history(self):
        s = self.store
        self.record()
        s.update_node(self.unit, "통신장치", "001", "시험용", "개별 설명", "#보관")
        nested = s.add_node("group", "하위 그룹", self.group)
        cable = s.add_node("device", "케이블", nested, description="신호 케이블")
        collection = s.save_collection("기존 모음", [self.unit])
        count = len(s.nodes())
        copy = s.duplicate_nodes([self.group, self.device])[0]
        self.assertEqual(s.node(copy)["name"], "장비 (1)")
        self.assertEqual(len(s.nodes()), count * 2)
        copied_nodes = [s.node(ident) for ident in s.descendants([copy])]
        copied_unit = next(n for n in copied_nodes if n["kind"] == "unit")
        self.assertEqual(copied_unit["serial"], "001")
        self.assertEqual(copied_unit["purpose"], "시험용")
        self.assertEqual(copied_unit["description"], "개별 설명")
        self.assertEqual(copied_unit["tags"], "#보관")
        self.assertEqual(s.entries(copied_unit["id"]), [])
        self.assertEqual(s.collection_targets(collection), [self.unit])
        second = s.duplicate_nodes([copy])[0]
        self.assertEqual(s.node(second)["name"], "장비 (2)")
        s.add_node("device", "통신장치 (1)", self.group)
        device_copy = s.duplicate_nodes([self.device])[0]
        self.assertEqual(s.node(device_copy)["name"], "통신장치 (2)")
        self.assertEqual(s.node(self.unit)["parent"], self.device)

    def test_duplicate_failure_rolls_back_subtree(self):
        s = self.store
        before = s.nodes()
        add = s.add_node
        calls = []
        def fail_second(*args, **kwargs):
            calls.append(True)
            if len(calls) == 2: raise ValueError("복제 중 오류")
            return add(*args, **kwargs)
        with patch.object(s, "add_node", fail_second):
            with self.assertRaises(ValueError): s.duplicate_nodes([self.group])
        self.assertEqual(s.nodes(), before)
        with self.assertRaises(ValueError): s.duplicate_nodes([self.unit])

    def test_serial_ranges(self):
        self.assertEqual(serial_range("001", "003"), ["001", "002", "003"])
        self.assertEqual(serial_range("P-09", "P-11"), ["P-09", "P-10", "P-11"])
        self.assertEqual(serial_range("9", "11"), ["9", "10", "11"])
        self.assertEqual(serial_range("SN-A", ""), ["SN-A"])
        self.assertEqual(serial_range("", ""), [])
        for a,b in [("", "1"), ("3", "1"), ("A1", "B2"), ("01", "100"), ("1", "10001")]:
            with self.assertRaises(ValueError): serial_range(a,b)

    def test_inherited_history_and_tags(self):
        parent = self.record(self.group, title="공통 점검", tag_text="#정기 #점검")
        direct = self.record(title="개별 점검")
        history = self.store.entries(self.unit)
        self.assertEqual({e["id"] for e in history}, {parent,direct})
        self.assertTrue(next(e for e in history if e["id"] == parent)["inherited"])
        self.assertEqual(len(self.store.entries(self.unit, inherited=False)), 1)
        self.assertEqual(len(self.store.entries(query="#점검")), 1)
        self.assertEqual(len(self.store.entries(query="#점")), 0)
        self.assertEqual(tags("현장 #현장, #시험"), "#현장 #시험")

    def test_import_transaction_rolls_back_all_rows(self):
        before = len(self.store.nodes())
        with self.assertRaises(sqlite3.IntegrityError):
            self.store.import_rows(self.group, [("새 장치", ["001"]), ("통신장치", ["002", "001"])])
        self.assertEqual(len(self.store.nodes()), before)

    def test_event_snapshot_deduplicates_and_unlinks(self):
        second = self.store.add_node("unit", "통신장치", self.device, "002")
        e = self.record(event=True, node_ids=[], category="출고", location="현장A")
        self.store.link_event(e, [self.group, self.device, self.unit])
        self.store.add_node("unit", "통신장치", self.device, "003")
        self.assertEqual(set(self.store.targets(e)), {self.unit,second})
        self.assertEqual(self.store.current_state(self.unit), ("현장A", "출고"))
        self.store.unlink_event(e, [self.unit])
        self.assertEqual(self.store.current_state(self.unit), ("미지정", "미기록"))
        self.assertEqual(self.store.node(self.unit)["serial"], "001")

    def test_state_uses_dates_and_ignores_future_and_group(self):
        self.record(category="입고", day="2026-02-01", location="창고")
        self.record(category="출고", day="2026-01-01", location="과거 현장")
        self.record(category="출고", day="2099-01-01", location="미래 현장")
        self.record(self.group, category="출고", day="2026-03-01", location="그룹 위치")
        self.assertEqual(self.store.current_state(self.unit), ("창고", "입고"))

    def test_move_and_cycle(self):
        nested = self.store.add_node("group", "하위", self.group)
        with self.assertRaises(ValueError): self.store.move_nodes([self.group], nested)
        self.store.move_nodes([self.device], nested)
        self.assertEqual(self.store.node(self.device)["parent"], nested)
        with self.assertRaises(ValueError): self.store.move_nodes([self.unit], nested)

    def test_delete_restore_parent_and_history(self):
        e = self.record()
        self.store.delete([self.group])
        self.assertEqual(self.store.nodes(), [])
        self.assertEqual(self.store.entries(), [])
        self.store.restore(self.store.trash()[0]["id"])
        self.assertEqual(len(self.store.nodes()), 3)
        self.assertEqual(self.store.entries(self.unit)[0]["id"], e)
        self.store.delete([e], "entries")
        self.assertEqual(self.store.entries(), [])
        self.store.restore(self.store.trash()[0]["id"])
        self.assertEqual(len(self.store.entries()), 1)

    def test_portability_and_backup(self):
        self.record(tag_text="#이동", location="창고")
        backup = self.store.backup()
        self.assertTrue(backup.is_file())
        with closing(sqlite3.connect(backup)) as db:
            self.assertEqual(db.execute("SELECT count(*) FROM entries").fetchone()[0], 1)
        target = Path(self.tmp.name) / "다른컴퓨터"
        shutil.copytree(self.store.root, target)
        moved = Store(target / "project.semp")
        try:
            self.assertEqual(moved.current_state(self.unit)[0], "창고")
            self.assertEqual(moved.entries()[0]["html"], "<b>메모</b>")
        finally: moved.close()

    def test_restore_conflict_is_atomic(self):
        self.store.delete([self.unit])
        token = self.store.trash()[0]["id"]
        self.store.add_node("unit", "통신장치", self.device, "001")
        with self.assertRaises(sqlite3.IntegrityError): self.store.restore(token)
        self.assertEqual(len(self.store.nodes()), 3)
        self.assertEqual(self.store.trash()[0]["id"], token)

    def test_edit_event_preserves_trashed_target_link(self):
        e = self.record(event=True, node_ids=[self.unit])
        self.store.delete([self.unit])
        token = self.store.trash()[0]["id"]
        self.record(event=True, node_ids=[], ident=e, title="수정된 이벤트")
        self.store.restore(token)
        self.assertEqual(self.store.targets(e), [self.unit])

    def test_editing_old_purpose_record_does_not_overwrite_current_purpose(self):
        entry = self.store.save_record(
            event=True, node_ids=[self.unit], title="환경용 변경", category="용도변경",
            day="2026-01-01", end_day="2026-01-01", html="환경용 변경",
            plain="환경용 변경", purpose_after="환경용")
        self.store.bulk_update_purpose([self.unit], "점검용")
        self.store.save_record(
            ident=entry, event=True, node_ids=[self.unit], title="본문 수정",
            category="용도변경", day="2026-01-01", end_day="2026-01-01",
            html="본문 수정", plain="본문 수정", purpose_after="환경용")
        self.assertEqual(self.store.node(self.unit)["purpose"], "점검용")

    def test_unit_search_uses_current_parent_device_name(self):
        entry = self.record(event=True)
        self.store.update_node(self.device, "통신장치-변경", "", "", "", "")
        self.assertEqual(self.store.node(self.unit)["name"], "통신장치-변경")
        self.assertEqual([record["id"] for record in self.store.entries(query="통신장치-변경")], [entry])


if __name__ == "__main__": unittest.main()
