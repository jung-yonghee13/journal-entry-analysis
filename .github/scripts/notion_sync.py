#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GitHub push → 노션 프로젝트 페이지 '업데이트 로그' 자동 기록.

GitHub Actions에서 실행된다. 표준 라이브러리만 사용(의존성 없음).
- NOTION_TOKEN     : 노션 내부 통합(integration) 토큰 (GitHub Secret)
- NOTION_PAGE_ID   : 갱신할 노션 페이지 ID (워크플로에 저장소별로 지정)
- COMMITS_JSON     : github.event.commits (push된 커밋 목록)
- HEAD_SHA         : push 후 HEAD SHA

동작: 페이지에서 '업데이트 로그' 제목 아래의 표를 찾아 헤더 바로 다음(=맨 위)에
새 행(일자 | 변경 요약 | 커밋)을 넣는다. 섹션이 없으면 페이지 끝에 새로 만든다.
변경 요약은 커밋 제목들에서 만든다(머지 커밋 제외, conventional prefix 제거).
토큰이 없으면 아무것도 하지 않고 성공 종료한다(설정 전 단계에서 CI 실패 방지).
"""
import datetime
import json
import os
import sys
import urllib.error
import urllib.request

API = "https://api.notion.com/v1"
TOKEN = os.environ.get("NOTION_TOKEN", "").strip().strip("﻿").strip()
PAGE_ID = os.environ.get("NOTION_PAGE_ID", "").strip().strip("﻿").strip()
COMMITS = os.environ.get("COMMITS_JSON", "[]")
SHA = os.environ.get("HEAD_SHA", "")[:7]

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}


def req(method, path, payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    r = urllib.request.Request(API + path, data=data, headers=HEADERS, method=method)
    with urllib.request.urlopen(r) as resp:
        return json.load(resp)


def children(block_id):
    out, cursor = [], None
    while True:
        q = "?page_size=100" + (f"&start_cursor={cursor}" if cursor else "")
        d = req("GET", f"/blocks/{block_id}/children{q}")
        out += d["results"]
        if not d.get("has_more"):
            return out
        cursor = d["next_cursor"]


def plain(block):
    body = block.get(block["type"], {})
    return "".join(rt.get("plain_text", "") for rt in body.get("rich_text", []))


def build_summary():
    try:
        commits = json.loads(COMMITS)
    except (ValueError, TypeError):
        commits = []
    subjects = []
    for c in commits:
        s = (c.get("message") or "").splitlines()[0].strip()
        if not s or s.startswith("Merge"):
            continue
        low = s.lower()
        for p in ("feat:", "fix:", "docs:", "chore:", "refactor:", "test:", "style:", "perf:"):
            if low.startswith(p):
                s = s[len(p):].strip()
                break
        subjects.append(s)
    if not subjects:
        subjects = ["커밋 반영"]
    text = " · ".join(subjects[:3])
    if len(subjects) > 3:
        text += f" 외 {len(subjects) - 3}건"
    return text[:180]


def text_cell(content, code=False):
    cell = {"type": "text", "text": {"content": content}}
    if code:
        cell["annotations"] = {"code": True}
    return [cell]


def data_row(date_s, summary, sha):
    return {"type": "table_row", "table_row": {
        "cells": [text_cell(date_s), text_cell(summary), text_cell(sha, code=True)]}}


def header_row():
    return {"type": "table_row", "table_row": {
        "cells": [text_cell("일자"), text_cell("변경 요약"), text_cell("커밋")]}}


def main():
    # 커밋 메시지에 [skip notion]이 있으면 기록하지 않는다 (CI 설정 등 메타 커밋용)
    try:
        msgs = [c.get("message", "") for c in json.loads(COMMITS)]
    except (ValueError, TypeError):
        msgs = []
    if msgs and all("[skip notion]" in m for m in msgs):
        print("[skip notion] 마커 — 노션 기록 건너뜀")
        return 0
    if not TOKEN or not PAGE_ID:
        print("NOTION_TOKEN 또는 NOTION_PAGE_ID 미설정 — 동기화 건너뜀 (Secret 설정 후 자동 활성화)")
        return 0
    row = data_row(datetime.date.today().isoformat(), build_summary(), SHA)

    blocks = children(PAGE_ID)
    table_id = None
    for i, b in enumerate(blocks):
        if b["type"].startswith("heading") and "업데이트 로그" in plain(b):
            for nb in blocks[i + 1:i + 4]:  # 제목 바로 뒤 인접 블록에서 표 탐색
                if nb["type"] == "table":
                    table_id = nb["id"]
                    break
            break

    if table_id:
        rows = children(table_id)
        payload = {"children": [row]}
        if rows:  # 헤더 행 바로 다음 = 목록 맨 위
            payload["after"] = rows[0]["id"]
        req("PATCH", f"/blocks/{table_id}/children", payload)
        print("업데이트 로그 표에 새 행 추가 완료")
    else:
        req("PATCH", f"/blocks/{PAGE_ID}/children", {"children": [
            {"type": "heading_1", "heading_1": {"rich_text": text_cell("🔄 업데이트 로그")}},
            {"type": "table", "table": {
                "table_width": 3, "has_column_header": True,
                "children": [header_row(), row]}},
        ]})
        print("업데이트 로그 섹션이 없어 페이지 끝에 신설 + 행 추가 완료")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")[:300]
        print(f"Notion API 오류 {e.code}: {body}", file=sys.stderr)
        sys.exit(1)
