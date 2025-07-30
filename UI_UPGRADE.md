# UI 업그레이드 가이드

## 변경 사항

기존 Tkinter 기반 UI를 CustomTkinter를 사용한 현대적인 Windows 11 스타일로 업그레이드했습니다.

### 주요 개선 사항

1. **Windows 11 디자인 언어 적용**
   - 둥근 모서리와 카드 기반 레이아웃
   - Windows 11 색상 팔레트 (Primary Blue: #0078D4)
   - Segoe UI 폰트 패밀리 사용

2. **카드 기반 레이아웃**
   - 헤더 카드: 타이틀과 설명
   - 점검 항목 카드: 체크박스로 진행 상황 표시
   - 진행 상태 카드: 실시간 상태 업데이트
   - 경고 카드: 주의사항 안내

3. **개선된 오버레이**
   - 둥근 모서리와 테두리 효과
   - 더 나은 가독성과 시각적 피드백
   - 부드러운 프로그레스바 애니메이션

4. **향상된 다이얼로그**
   - CustomTkinter 스타일의 확인 대화상자
   - 결과 표시용 커스텀 다이얼로그
   - 통계와 결과를 시각적으로 표시

## 설치 방법

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

2. 프로그램 실행:
```bash
python main.py
```

## 시스템 요구사항

- Windows 10/11
- Python 3.8 이상
- CustomTkinter 5.2.1
- 기타 기존 의존성 (pyautogui, pygetwindow 등)

## 주요 기능 (변경 없음)

모든 기존 기능은 그대로 유지됩니다:
- 10개의 보안 점검 작업
- 자동 스크린샷 캡처
- Word 문서 생성
- 작업 중 오버레이 표시

## 색상 테마

프로그램은 라이트 모드로 설정되어 있으며, 다음 색상을 사용합니다:
- Primary: #0078D4 (Windows Blue)
- Success: #107C10 (Green)
- Warning: #FFB900 (Yellow)
- Danger: #D83B01 (Red)
- Card Background: #FFFFFF
- Text Primary: #323130
- Text Secondary: #605E5C

다크 모드를 원하시면 main.py의 872번째 줄을 다음과 같이 변경하세요:
```python
ctk.set_appearance_mode("dark")  # "light" 대신 "dark"
```