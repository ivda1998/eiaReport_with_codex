# 참고문헌 관리

이 폴더는 학술논문 검색 결과와 인용 정보를 저장합니다.

## 파일 구조
- `papers.jsonl`: 논문 메타데이터(JSONL). 각 줄이 하나의 논문입니다.
- `citations.bib`: BibTeX 형식 인용 파일.

## 사용 예시
```bash
python t2.py search "AERMOD 대기질" --rows 10
```

검색 결과는 `papers.jsonl`과 `citations.bib`에 누적됩니다. 보고서 생성 시 자동으로 참고문헌 섹션에 포함됩니다.
