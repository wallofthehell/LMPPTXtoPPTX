# LMPPTXtoPPTX

NotebookLM(Gemini)에서 내려받은 PPTX는 슬라이드 전체가 **통그림 1장**이라 글자를 편집할 수 없습니다.
이 도구는 글자를 인식(OCR) → 사람이 확인·수정 → 그림에서 글자를 지우고 **편집 가능한 텍스트 상자**로 올립니다.
로고·아이콘·카드·사진 등 디자인은 배경 그림으로 그대로 남고, 슬라이드 복제/삭제/이동은 PowerPoint 기본 기능으로 하면 됩니다.

> OCR은 한글+영문 혼합 글자를 틀릴 수 있습니다(예: `AI로` → `시로`). 그래서 변환 전 확인 단계가 필수입니다.

## 웹 (설치 불필요)
https://wallofthehell.github.io/LMPPTXtoPPTX/ — 브라우저에서만 처리되며 파일이 서버로 전송되지 않습니다.
1. PPTX 선택 → **글자 인식** 2. 원본 글자 조각과 대조해 수정, 잘못 잡힌 줄(아이콘 등)은 체크 해제 3. **변환하고 다운로드**

## Python CLI (로컬/오프라인)
```
pip install -r requirements.txt     # + Tesseract 설치 및 kor 언어 데이터
python lmpptx2pptx.py analyze in.pptx draft.json      # draft.json + draft_slideN.png(번호 박스 미리보기)
#  draft.json 의 text 수정, 불필요한 줄은 "skip": true
python lmpptx2pptx.py build in.pptx draft.json out.pptx --font "맑은 고딕"
```

## GitHub Pages 배포
Settings → Pages → Source: *Deploy from a branch* → `main` / `(root)` → Save. 1~2분 뒤 위 주소에서 열립니다.

## 한계
- 슬라이드 전체 크기의 그림이 있는 슬라이드만 처리합니다(이미 텍스트 개체가 있으면 건드리지 않음).
- 글자 크기·굵기·색은 이미지에서 추정하므로 PowerPoint에서 미세 조정이 필요할 수 있습니다.
- 글머리표(•)·아이콘은 배경 그림에 남습니다. 지운 자리는 단순 보간/인페인팅이라 복잡한 배경에서는 흔적이 남을 수 있습니다.
