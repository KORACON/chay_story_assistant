import json
from pathlib import Path
from typing import List


class AccessManager:
    def __init__(self, file_path: Path, admin_ids: List[int]):
        self.file_path = file_path
        self.admin_ids = set(admin_ids)
        self.allowed_ids = set()
        self._load()

    def _load(self):
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.file_path.exists():
            self._save()
            return
        try:
            data = json.loads(self.file_path.read_text(encoding="utf-8"))
            self.allowed_ids = set(int(x) for x in data.get("allowed_ids", []))
        except Exception:
            self.allowed_ids = set()
            self._save()

    def _save(self):
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        data = {"allowed_ids": sorted(self.allowed_ids)}
        self.file_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def is_admin(self, user_id: int) -> bool:
        return user_id in self.admin_ids

    def is_allowed(self, user_id: int) -> bool:
        return self.is_admin(user_id) or (user_id in self.allowed_ids)

    def add_user(self, user_id: int):
        if user_id not in self.admin_ids:
            self.allowed_ids.add(user_id)
            self._save()

    def del_user(self, user_id: int):
        if user_id in self.allowed_ids:
            self.allowed_ids.remove(user_id)
            self._save()

    def list_users(self) -> List[int]:
        return sorted(self.allowed_ids)
