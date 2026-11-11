# -*- coding: utf-8 -*-
"""R11 覆盖充分性核验回归测试：固定 20 条公开标准判例，结论不得漂移。

    python tests/test_coverage_cases.py

样本来自真实使用记录，条款全部取自公开标准（GB/DL/JGJ/AQ），
问题已脱敏为通用场景，见 tests/fixtures/coverage_cases.json。

改了 server/private/coverage_core.py 的关键词、阈值或维度判定后跑一遍，
PASS 数量下降说明有回归；确实是有意调整时，用 `--update` 重算期望值并复核 diff。
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "server" / "private"))

os.environ.pop("ANTHROPIC_API_KEY", None)  # 回归测试只验证启发式，不触发联网复核

from coverage_core import check_coverage  # noqa: E402

FIXTURE = REPO / "tests" / "fixtures" / "coverage_cases.json"

ok = 0
fail = 0


def check(name: str, cond: bool, detail: str = "") -> None:
    global ok, fail
    if cond:
        ok += 1
        print(f"  PASS  {name}")
    else:
        fail += 1
        print(f"  FAIL  {name} {detail}")


def load_cases() -> list[dict]:
    data = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return data["cases"]


def test_fixture(cases: list[dict]) -> None:
    print("[夹具结构]")
    check("样本数 >= 20", len(cases) >= 20, f"实际 {len(cases)}")
    ids = [c["id"] for c in cases]
    check("id 唯一", len(ids) == len(set(ids)))
    empty = [c["id"] for c in cases if not c.get("clause", "").strip()]
    check("条款非空", not empty, str(empty))
    missing = [c["id"] for c in cases if not {"id", "question", "clause", "covered"} <= set(c)]
    check("必填字段齐全", not missing, str(missing))


def test_cases(cases: list[dict], update: bool) -> None:
    print("[覆盖核验回归]")
    log_path = os.path.join(tempfile.mkdtemp(), "judgment_log.jsonl")
    changed = []
    for case in cases:
        result = check_coverage(case["question"], case["clause"], log_path=log_path)
        same = result["covered"] is case["covered"]
        check(
            f"{case['id']} {case['source']} 期望 covered={case['covered']}",
            same,
            f"实际 covered={result['covered']}",
        )
        if not same:
            changed.append((case, result))
    if update and changed:
        for case, result in changed:
            case["covered"] = result["covered"]
            case["dimensions_passed"] = result.get("dimensions_passed", [])
            case["note"] = result.get("reason_summary", "")
        data = json.loads(FIXTURE.read_text(encoding="utf-8"))
        data["cases"] = cases
        FIXTURE.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"\n已更新 {len(changed)} 条期望值，请复核 git diff 后再提交")


def main() -> int:
    cases = load_cases()
    test_fixture(cases)
    test_cases(cases, update="--update" in sys.argv)
    print(f"\n结果：PASS {ok} / FAIL {fail}")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
