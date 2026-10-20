# PDF 편집기 4단계 — 양식 채우기·서명 실행 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** PDF 양식 칸을 편집기에서 채우고(필드 창 포함), 손글씨·그림·글자 서명과 날짜·글자 도장을 쪽에 놓아 저장 파일에 표준으로 남긴다.

**Architecture:** 양식 값은 방의 `pdfForm`(칸 이름 → 값)에 두고, 쪽마다 pdf.js가 주는 위젯 정보로 우리 입력란을 그린다(지금은 ENABLE_FORMS라 본문에 양식이 아예 안 보이던 것도 해결). 저장 때 pdf-lib로 채우고 한글은 Pretendard로 모양을 만든다. 서명·도장은 새 주석 종류 `stamp`(그림 data URL) — 표준 Stamp 주석 + 그림 사본(`/DocviewImage`)으로 써서 다시 열어도 고칠 수 있다.

**Tech Stack:** pdf-lib(form·embedPng/Jpg), pdfjs-dist(getAnnotations 위젯), Yjs, Playwright.

**Spec:** docs/superpowers/specs/2026-10-20-pdf-editor-design.md §1(서명 탭·필드 창), §2.2(pdfForm), §2.3-3·4, §4-4

## Global Constraints

- 8000 영향 없음(pdfEditor 토큰만), 1~3단계 기능·다른 편집기 회귀 없음(npm test, verify:xlsx/pptx/pdf/pdf-annots/pdf-organize, 워드 리본).
- 양식 값은 바꾼 칸만 쓴다. 읽기 전용 칸·읽기 전용 토큰은 입력 불가.
- 서명 그림은 1.5MB(data URL) 이하, 편집기에서 긴 변 800px로 줄인다.

---

### Task 1: 서버 — stamp 종류, 양식 채우기, 지운 쪽의 양식 칸 정리, 변경 기록 (tests/pdfForm.test.ts)
### Task 2: 편집기 양식 층(formLayer.ts) — 글자·체크·라디오·목록 입력란, 같은 이름 칸 동기화, 필드 창, 실행 취소
### Task 3: 서명 탭 — 서명 만들기 창(그리기·그림·글자), 내 서명 목록(브라우저 저장), 놓기·옮기기·크기(비율 유지), 날짜·글자 도장
### Task 4: e2e(verify:pdf-forms) → 회귀 → 반영 → 운영 확인
