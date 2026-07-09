# sample_data — 예시 분개장 데이터

플러그인 시연·테스트용 예시 파일 폴더입니다. **제출 양식(src/, README.md, logs/)과는 별개**이며,
분석하고 싶은 분개장 파일을 이 폴더에 넣고 사용하면 됩니다.

| 파일 | 내용 |
|---|---|
| `예시분개장1.xlsx` | 2025 회계연도 가상 분개장 (전표 2,025매 / 4,892행, 이상 전표 27매 포함) |
| `정답지.xlsx` | 위 파일에 심어진 이상 전표 27매의 목록 — 탐지 성능 검증용 |

## 사용 예

Codex/Claude에서:
```
sample_data/예시분개장1.xlsx 분개장 검토해줘
```

또는 스크립트 직접 실행:
```
python src/skills/journal-entry-analysis/scripts/analyze_journal.py --input sample_data/예시분개장1.xlsx
```

> 참고: 분석을 실행하면 입력 파일에 `이상전표` 시트가 추가되고(추가 전 자동 백업),
> `<파일명>_분석/` 폴더에 분석 보고서(PDF·Word·Markdown)와 차트가 생성됩니다.
> 탐지 성능: 정답지 27/27 (검증 완료)
