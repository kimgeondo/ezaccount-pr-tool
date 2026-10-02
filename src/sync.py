import json
from datetime import datetime, timezone
from pathlib import Path


def load_state(path):
    state_path = Path(path)
    if not state_path.exists():
        return {"last_sync": None}
    return json.loads(state_path.read_text(encoding="utf-8"))


def save_state(path, state):
    state_path = Path(path)
    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def get_last_sync(state):
    return state.get("last_sync")


def build_state_update():
    return {"last_sync": datetime.now(timezone.utc).replace(microsecond=0).isoformat()}


def item_key(item):
    pr_no = item.get("PRNo") or item.get("PR No") or item.get("DocumentNumber") or item.get("Document Number")
    item_no = item.get("ItemNo") or item.get("Item No") or item.get("LineItem")
    if pr_no is None:
        pr_no = ""
    if item_no:
        return f"{pr_no}|{item_no}"
    return "|".join(
        [
            pr_no,
            str(item.get("ShortText", "") or ""),
            str(item.get("Quantity", "") or ""),
            str(item.get("VendorName", "") or ""),
        ]
    )


def dedupe_items(items):
    seen = set()
    cleaned = []
    for item in items:
        key = item_key(item)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(item)
    return cleaned
