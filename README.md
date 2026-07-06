# 늘봄학교 합성 사용자 FGI 시스템

## 실행 방법 (로컬)
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 배포 방법 (Streamlit Cloud)
1. GitHub에 이 폴더 전체 업로드
2. https://share.streamlit.io 접속
3. New app → GitHub 레포 선택 → app.py 선택
4. Secrets 설정 불필요 (API 키는 앱 내에서 입력)
5. Deploy!

## 폴더 구조
```
neulbom_fgi/
├── app.py              # 메인 앱
├── requirements.txt    # 패키지 목록
├── data/               # RAG 데이터 파일
│   ├── *.pdf
│   └── *.xlsx
└── chroma_db/          # 벡터DB (자동 생성, GitHub 제외)
```

## 사용 방법
1. 왼쪽 사이드바에 Anthropic API 키 입력
2. 정책 단계 선택 (탐색 / 검증)
3. 질문 입력 → 합성 이해관계자 3인 순차 응답
4. 세션 종료 후 대화록 저장
