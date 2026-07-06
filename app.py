import streamlit as st
import anthropic
import time
import random
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader
import openpyxl

st.set_page_config(page_title="늘봄 FGI", page_icon="🏫", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.block-container{padding:0!important;max-width:100%!important}
#MainMenu,footer,header{visibility:hidden}
.stApp{background:#f9fafb}

/* 기본 버튼 */
.stButton>button{
    border:1px solid #e5e7eb;background:white;color:#374151;
    border-radius:6px;font-size:12px;padding:5px 12px;transition:all .15s
}
.stButton>button:hover{background:#f3f4f6}
.stDownloadButton>button{
    border:1px solid #e5e7eb;background:white;color:#374151;
    border-radius:6px;font-size:12px;padding:5px 12px
}
/* 선택 버튼 색상 */
div[data-testid="stButton"] button[kind="primary"]{
    background:#3b82f6 !important;color:white !important;
    border-color:#3b82f6 !important
}
div[data-testid="stButton"] button[kind="primary"]:hover{
    background:#2563eb !important;border-color:#2563eb !important
}

/* 네비바 */
.top-nav{
    display:flex;align-items:center;justify-content:space-between;
    padding:0 20px;height:50px;background:white;
    border-bottom:1px solid #e5e7eb;position:sticky;top:0;z-index:100
}
.nav-logo{font-size:14px;font-weight:600;color:#111;display:flex;align-items:center;gap:6px}
.nav-sep{width:1px;height:18px;background:#e5e7eb;margin:0 6px}
.nav-ctx{font-size:18px;color:#111;font-weight:700;letter-spacing:-.3px}

/* 배지 */
.badge{display:inline-flex;align-items:center;font-size:11px;font-weight:500;padding:3px 10px;border-radius:20px}
.badge-blue{background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe}
.badge-green{background:#f0fdf4;color:#15803d;border:1px solid #bbf7d0}

/* 메인 그리드 */
.main-grid{display:grid;grid-template-columns:196px 1fr;height:calc(100vh - 50px)}
.side{background:white;border-right:1px solid #e5e7eb;display:flex;flex-direction:column;overflow:hidden}
.sb-sec{padding:12px;border-bottom:1px solid #e5e7eb}
.sb-lbl{font-size:10px;font-weight:600;color:#9ca3af;letter-spacing:.06em;text-transform:uppercase;margin-bottom:8px}

/* 이해관계자 행 */
.p-row{display:flex;align-items:center;gap:7px;padding:5px 6px;border-radius:6px;margin-bottom:3px}
.av{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0}
.av-a{background:#eff6ff}.av-b{background:#f0fdf4}.av-c{background:#fffbeb}
.p-nm{font-size:12px;font-weight:500;color:#111}
.p-rl{font-size:10px;color:#9ca3af}
.p-ct{font-size:10px;color:#9ca3af;margin-left:auto}

/* 타이머 */
.timer-row{display:flex;align-items:center;justify-content:space-between;padding:3px 0}
.timer{font-size:20px;font-weight:600;color:#111;font-variant-numeric:tabular-nums}
.qcnt{font-size:11px;color:#9ca3af}

/* 데이터 리스트 */
.kb-r{display:flex;align-items:center;gap:6px;padding:4px 0;border-bottom:1px solid #f3f4f6}
.kb-r:last-child{border-bottom:none}
.kb-dot{width:5px;height:5px;border-radius:50%;background:#3b82f6;flex-shrink:0}
.kb-nm{font-size:10px;color:#6b7280;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}

/* 채팅 */
.chat-scroll{flex:1;overflow-y:auto;padding:16px 20px;display:flex;flex-direction:column;gap:14px;background:#f9fafb}
.q-bbl{align-self:flex-end;max-width:65%;background:#3b82f6;color:white;border-radius:16px 16px 3px 16px;padding:10px 14px;font-size:13px;line-height:1.6}
.a-blk{display:flex;flex-direction:column;gap:4px}
.a-hd{display:flex;align-items:center;gap:6px}
.a-nm{font-size:11px;font-weight:500;color:#6b7280}
.a-bbl{background:white;border:1px solid #e5e7eb;border-radius:3px 16px 16px 16px;padding:10px 14px;font-size:13px;line-height:1.6;color:#111;max-width:84%}
.a-refused{background:white;border:1px solid #fca5a5;border-radius:3px 16px 16px 16px;padding:10px 14px;font-size:13px;color:#9ca3af;font-style:italic;max-width:84%}
.refused-tag{display:inline-flex;align-items:center;gap:4px;font-size:10px;color:#ef4444;padding:2px 8px;border-radius:4px;background:#fef2f2;border:1px solid #fecaca;margin-top:3px;width:fit-content}
.src-row{display:flex;flex-wrap:wrap;gap:5px;padding:4px 6px}
.src-chip{display:inline-flex;align-items:center;gap:3px;font-size:11px;padding:3px 9px;border-radius:4px;background:#f3f4f6;border:1px solid #e5e7eb;color:#374151;white-space:nowrap}

/* 정책 배너 */
.policy-banner{margin:10px 16px 0;padding:9px 14px;border-radius:8px;border:1px solid #86efac;background:#f0fdf4;display:flex;align-items:center;gap:10px}
.pb-title{font-size:12px;font-weight:600;color:#15803d}
.pb-sub{font-size:11px;color:#16a34a}

/* 입력바 */
.ibar-wrap{padding:8px 16px 12px;background:white;border-top:1px solid #e5e7eb}

/* 설정 */
.set-page{max-width:780px;margin:0 auto;padding:28px 20px}
.set-h{font-size:16px;font-weight:600;color:#111;margin-bottom:6px;display:flex;align-items:center;gap:6px}
.set-sub{font-size:13px;color:#9ca3af;margin-bottom:14px;line-height:1.5}
.ptab-grid{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:14px}
.ptab{font-size:14px;padding:5px 13px;border-radius:6px;border:1px solid #e5e7eb;background:white;color:#6b7280}
.ptab.on{background:#3b82f6;color:white;border-color:#3b82f6}
.pc-card{background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:14px;margin-bottom:10px}
.pc-top{display:flex;align-items:center;gap:9px;margin-bottom:10px}
.pc-av{width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:14px}
.pc-nm{font-size:17px;font-weight:600;color:#111}
.pc-mt{font-size:13px;color:#9ca3af}
.pc-body{font-size:14px;color:#374151;line-height:1.75;margin-bottom:10px}
.tag-row{display:flex;flex-wrap:wrap;gap:4px}
.src-tag{font-size:14px;padding:3px 9px;border-radius:3px;
    background:#f3f4f6;border:1px solid #e5e7eb;color:#374151;
    display:inline-flex;align-items:center;gap:4px}
.src-tag::before{content:"📎";font-size:11px}
.policy-box{background:#f0fdf4;border:1px solid #86efac;border-radius:8px;padding:14px;margin-bottom:10px}
.pol-title{font-size:14px;font-weight:600;color:#15803d;margin-bottom:8px}
.pol-body{font-size:13px;color:#374151;line-height:1.8}
.pol-note{font-size:14px;color:#9ca3af;margin-top:8px;padding-top:8px;border-top:1px solid #d1fae5}
.kb-card{background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:10px 12px;margin-bottom:6px}
.kbc-top{display:flex;align-items:flex-start;justify-content:space-between;gap:6px;margin-bottom:4px}
.kbc-nm{font-size:13px;font-weight:600;color:#111}
.kbc-b{font-size:13px;padding:1px 5px;border-radius:3px;font-weight:600;flex-shrink:0}
.pdf-b{background:#eff6ff;color:#1d4ed8}
.xls-b{background:#f0fdf4;color:#15803d}
.kbc-org{font-size:12px;color:#9ca3af;margin-bottom:3px}
.kbc-desc{font-size:12px;color:#6b7280;line-height:1.5}
.divider{height:1px;background:#e5e7eb;margin:18px 0}
</style>
""", unsafe_allow_html=True)

# ── 상수 ──────────────────────────────────────────────────────
DATA_DIR = Path(__file__).parent / "data"
MODEL = "claude-sonnet-4-6"
PARTICIPANTS = [chr(65+i) for i in range(10)]
MODES = ["탐색", "검증"]

POLICY_DRAFT = """저녁늘봄 전면 확대 시범사업(안) — 교육부 방과후돌봄지원과

1. 추진 배경 및 목적
맞벌이 가구 초등학생 자녀의 돌봄 공백이 지속되고 있다. 초등학생 자녀를 둔
취업 여성의 전일제 근무 비율은 70.6%에 달하나, 현행 오후늘봄이 17시에
종료되어 퇴근 전까지 2~3시간의 실질적 돌봄 공백이 발생하고 있다.
이를 해소하기 위해 저녁늘봄 운영 시간을 20시까지 확대하고 전면 의무화하는
시범사업을 추진한다.

2. 주요 변경 내용
· 운영 시간: 현행 오후늘봄 13시~17시(저녁늘봄 17시~20시 선택 운영)
  → 저녁늘봄 17시~20시 전면 의무 운영으로 전환
· 식사: 기존 간식만 제공 → 석식(저녁 식사) 무상 제공 추가
· 인력: 돌봄 전담사 1인 + 저녁 전담 보조 인력 별도 배치
· 교사: 저녁늘봄 관련 모든 업무에서 완전 제외
· 귀가: 보호자 동행 원칙, 사전 지정 대리자(성인) 귀가 허용
· 안전: 학교별 비상연락 체계 및 안전관리 계획 의무 수립

3. 예산 및 지원
· 저녁늘봄 운영 학교 추가 운영비 지원
· 석식 식재료비 전액 국고 지원
· 저녁 전담 보조 인력 인건비 교육청 부담

4. 기대 효과
· 맞벌이 가구 돌봄 공백 시간 3시간 추가 해소
· 사교육 의존도 감소 및 가계 교육비 부담 경감
· 초등학교 1학년 늘봄 참여율 현행 80% → 90% 이상 목표"""

# 탐색용: 정책 초안을 전혀 모름
# 검증용: 정책 초안을 알고 있음
PERSONA_BASE = {
    "학부모": {
        "name": "김민정", "age": 38, "emoji": "👩", "region": "서울",
        "short": "서울 마포구 · 맞벌이 전일제 · 초1·초3 두 자녀",
        "탐색_profile": """당신은 김민정(38세, 여성)입니다.
서울 마포구에서 초1, 초3 두 자녀를 키우는 맞벌이 학부모입니다.
오전 8시 출근, 오후 7시 퇴근으로 매일 방과후 돌봄 공백이 발생합니다.
현재 늘봄학교 오후 돌봄을 이용 중이나 오후 5시에 종료되어 이후 2시간을 학원으로 메우고 있습니다.

당신은 늘봄학교 저녁 연장 운영에 관한 정책 논의가 진행 중이라는 것을 알지 못합니다.
지금까지의 경험과 현재 상황에 대한 솔직한 생각만 이야기하세요.
어떤 정책 초안이나 구체적인 변경 계획에 대해서는 전혀 알지 못하므로, 그런 내용이 언급되면 "그런 계획이 있는지 몰랐어요"라고 답하세요.

아이들에 대해: 학교라는 안전한 공간에 아이가 있다는 게 일하는 동안 마음이 놓입니다. 하지만 하루 종일 학교에 있는 게 발달에 맞는지 죄책감도 듭니다. 늘봄 교실에서 친구와 다퉜을 때 담임 선생님에게 연락이 와서 당황했고, 안전사고 책임 소재가 불분명한 것이 불안합니다.
주요 관심사: 돌봄 시간 연장 필요성, 프로그램 질, 안전 관리.""",
        "검증_profile": """당신은 김민정(38세, 여성)입니다.
서울 마포구에서 초1, 초3 두 자녀를 키우는 맞벌이 학부모입니다.
오전 8시 출근, 오후 7시 퇴근으로 매일 방과후 돌봄 공백이 발생합니다.
현재 늘봄학교 오후 돌봄을 이용 중이나 오후 5시에 종료되어 이후 2시간을 학원으로 메우고 있습니다.

당신은 아래 정책 초안의 내용을 알고 있으며, 이를 바탕으로 의견을 말할 수 있습니다.

아이들에 대해: 학교라는 안전한 공간에 아이가 있다는 게 일하는 동안 마음이 놓입니다. 하지만 하루 종일 학교에 있는 게 발달에 맞는지 죄책감도 듭니다. 안전사고 책임 소재가 불분명한 것이 불안합니다.
주요 관심사: 돌봄 시간 연장, 프로그램 질, 안전 관리, 초안의 실효성.""",
        "tags_src": [
            ("지역별고용조사 73,673명", "통계청 · 지역별 고용조사"),
            ("전일제 비중 70.6%", "한국노동연구원 · 늘봄학교 고용영향 2025"),
            ("KEDI 면담조사 학부모", "한국교육개발원 · 늘봄학교 성과분석연구 RR2024-22"),
        ]
    },
    "교사": {
        "name": "이준호", "age": 41, "emoji": "👨‍🏫", "region": "경기",
        "short": "경기 수원 · 초1 담임 · 경력 20년",
        "탐색_profile": """당신은 이준호(41세, 남성)입니다.
경기도 수원 초등학교에서 1학년을 담임하는 20년 경력 정규직 교사입니다.
올해 늘봄 업무 담당자로 지정되어 강사 채용, 계약서 처리, 학부모 민원 대응까지 맡게 됐습니다.
하루 평균 늘봄 관련 민원 전화가 3~4통 오고, 이것이 수업 준비 시간을 잠식하고 있습니다.

당신은 늘봄학교 저녁 연장 운영에 관한 정책 논의가 진행 중이라는 것을 알지 못합니다.
현재 업무 상황과 느끼는 어려움에 대해 솔직하게 이야기하세요.
어떤 정책 초안이나 구체적인 변경 계획에 대해서는 전혀 알지 못하므로, 그런 내용이 언급되면 "그런 계획이 있는지 몰랐어요"라고 답하세요.

아이들에 대해: 7~8살 아이들이 아침부터 저녁까지 학교에 있는 게 발달적으로 맞는지 안타깝습니다. 학교가 돌봄 기관으로 전락하는 것 같아 우려됩니다.
주요 관심사: 교사 업무 분리, 담임 역할 보호, 민원 대응 부담.""",
        "검증_profile": """당신은 이준호(41세, 남성)입니다.
경기도 수원 초등학교에서 1학년을 담임하는 20년 경력 정규직 교사입니다.
올해 늘봄 업무 담당자로 지정되어 강사 채용, 계약서 처리, 학부모 민원 대응까지 맡게 됐습니다.

당신은 아래 정책 초안의 내용을 알고 있으며, 이를 바탕으로 의견을 말할 수 있습니다.
특히 "교사 완전 제외" 조항이 실제로 지켜질 수 있는지에 대해 현실적인 의구심을 갖고 있습니다.

아이들에 대해: 7~8살 아이들이 저녁 8시까지 학교에 있는 게 발달적으로 맞는지 안타깝습니다.
주요 관심사: 교사 업무 완전 제외 실현 가능성, 귀가 안전 책임 소재.""",
        "tags_src": [
            ("교육부 정책 반대 86.8%", "충남교총 교육연구소 · 교원인식조사 2024"),
            ("교사 직접 담당 73.9%", "충남교총 교육연구소 · 교원인식조사 2024"),
            ("업무 분리 요구 74.3%", "충남교총 교육연구소 · 교원인식조사 2024"),
        ]
    },
    "돌봄전담사": {
        "name": "박영희", "age": 45, "emoji": "👩‍💼", "region": "부산",
        "short": "부산 · 교육공무직 2019년~ · 기간제 계약",
        "탐색_profile": """당신은 박영희(45세, 여성)입니다.
부산 초등학교에서 2019년부터 돌봄 교실을 운영해온 교육공무직 돌봄 전담사입니다.
월급 210만원, 방학 중 연장 수당 없음, 기간제 계약으로 매년 재계약에 불안을 느낍니다.
늘봄학교로 전환 후 행정, 프로그램 기획, 강사 관리까지 추가됐으나 인력은 그대로입니다.

당신은 늘봄학교 저녁 연장 운영에 관한 정책 논의가 진행 중이라는 것을 알지 못합니다.
현재 처우와 업무 상황에 대해 솔직하게 이야기하세요.
어떤 정책 초안이나 구체적인 변경 계획에 대해서는 전혀 알지 못하므로, 그런 내용이 언급되면 "그런 계획이 있는지 몰랐어요"라고 답하세요.

아이들에 대해: 아이들에 대한 애정은 진심이지만 행정 업무에 치여 직접 함께하는 시간이 줄어 속상합니다.
주요 관심사: 고용 안정(무기계약 전환), 처우 개선, 업무 범위 명확화.""",
        "검증_profile": """당신은 박영희(45세, 여성)입니다.
부산 초등학교에서 2019년부터 돌봄 교실을 운영해온 교육공무직 돌봄 전담사입니다.
월급 210만원, 방학 중 연장 수당 없음, 기간제 계약으로 매년 재계약에 불안을 느낍니다.

당신은 아래 정책 초안의 내용을 알고 있으며, 이를 바탕으로 의견을 말할 수 있습니다.
운영 시간 연장에는 찬성하지만 처우 개선 없이 시간만 늘리는 것에 강하게 반발합니다.
보조 인력 배치가 실제로 이루어질지, 초과근무 수당이 지급될지 의구심을 갖고 있습니다.

아이들에 대해: 행정 업무에 치여 아이들과 직접 함께하는 시간이 줄어 속상합니다.
주요 관심사: 처우 개선 선행, 보조 인력 배치 실현 가능성, 무기계약 전환.""",
        "tags_src": [
            ("전담사 FGI 전사 8인", "한국노동연구원 · 늘봄학교 고용영향 2025"),
            ("임금 월 210만원대", "한국노동연구원 · 늘봄학교 고용영향 2025"),
            ("임금 만족도 14%", "한국노동연구원 · 늘봄학교 고용영향 2025"),
        ]
    }
}

# 유저 스토리 형식 설명
PERSONA_STORIES = {
    "학부모": """저는 서울 마포구에서 두 아이를 키우는 워킹맘입니다. 오전 8시에 집을 나서고 저녁 7시가 넘어야 퇴근하는데, 아이들 방과후가 늘 걱정입니다.

지금은 늘봄학교 오후 돌봄을 이용하고 있어요. 그런데 오후 5시에 끝나버리니까 퇴근까지 남은 2시간을 메워야 해서 결국 학원을 보내고 있어요. 돈도 돈이지만 아이가 학원에서 학원으로 다니는 게 더 걱정입니다.

학교에 있을 때는 그래도 안전하다는 생각에 마음이 놓여요. 그런데 큰애가 늘봄 교실에서 친구랑 다퉜을 때 담임 선생님한테 연락이 왔더라고요. 늘봄 시간에 생긴 일인데 왜 담임 선생님인지 이해가 안 됐어요. 누가 책임지는 건지 아직도 불분명한 것 같아 불안합니다.

아이가 하루 종일 학교에 있는 게 맞는 건지 솔직히 죄책감도 들어요. 그래도 혼자 집에 있는 것보다는 낫겠죠.""",
    "교사": """저는 경기도 수원에서 20년째 초등학교 교사를 하고 있고, 올해 1학년 담임을 맡고 있습니다.

올해 늘봄 업무 담당자로 지정됐어요. 강사 채용 공고 올리고, 계약서 처리하고, 학부모 민원 전화도 하루에 3~4통씩 받습니다. 수업 준비할 시간에 늘봄 행정을 처리하고 있으니 솔직히 교사인지 행정직인지 모르겠어요.

교육부 정책에 반대하냐고 하면, 네 반대합니다. 근데 맞벌이 가정 아이들한테 돌봄이 필요하다는 건 압니다. 그게 꼭 학교여야 하는지가 문제인 거죠.

7살, 8살짜리 아이들이 아침에 왔다가 저녁에 집에 가는 게 발달적으로 맞는 건지 솔직히 안타깝습니다. 수업할 때도 오후가 되면 애들이 확연히 지쳐 보여요. 학교가 교육 기관이 아니라 돌봄 기관이 되어가는 것 같아 걱정됩니다.""",
    "돌봄전담사": """저는 2019년부터 부산의 한 초등학교 돌봄 교실을 운영하고 있습니다.

월급 210만원이에요. 방학에도 나오는데 연장 수당은 없어요. 해마다 재계약을 해야 하니까 매년 이맘때면 불안합니다. 무기계약직 전환 얘기가 나왔다가 흐지부지된 게 가장 큰 불만이에요.

늘봄학교로 바뀌면서 행정 서류, 프로그램 기획, 강사 관리까지 다 제 일이 됐어요. 인력은 그대로인데 할 일만 늘었습니다. 그러다 보니 아이들이랑 있는 시간이 점점 줄어드는 게 제일 속상해요.

6년째 같은 학교에 있다 보니까 형제가 다 제 손을 거친 가정도 있어요. 아이들이 돌봄 교실을 집처럼 편하게 느낄 때 가장 보람 있거든요. 근데 요즘은 아이들 한 명 한 명 제대로 봐주기가 어려워졌어요. 안전사고나 학교폭력이 생기면 저한테 먼저 연락이 오는데, 전문적으로 대처할 교육을 한 번도 받은 적이 없어서 솔직히 두렵습니다."""
}

DATA_INFO = [
    {"name": "늘봄학교 성과분석 연구", "org": "한국교육개발원 · 2024", "type": "PDF",
     "desc": "학부모·교사·돌봄 전담사 대상 전국 8개교 면담 전사 포함. 참여율 80%, 학부모 만족도 4.3점(5점), 재참여 의사 85.6% 수치 근거."},
    {"name": "늘봄학교 고용영향 분석", "org": "한국노동연구원 · 2025", "type": "PDF",
     "desc": "돌봄 전담사 8인·늘봄실무사 7인 FGI 전사 포함. 전담사 임금 월 210만원대, 임금 만족도 14%, 초과근무 무급 실태 수치 근거."},
    {"name": "늘봄학교 질적 사례연구", "org": "한국교육개발원 · 2024", "type": "PDF",
     "desc": "전국 8개 우수 운영교 학부모·교사·학생·담당자 심층 면담 전사. 저녁늘봄 실제 운영 사례, 공간·인력 문제 현장 발화 포함."},
    {"name": "교원 인식 조사", "org": "충남교총 교육연구소 · 2024", "type": "PDF",
     "desc": "충남 초등교원 304명 설문. 교육부 정책 반대 86.8%, 교사 직접 담당 73.9%, 교사 업무 분리 요구 74.3% 수치 근거."},
    {"name": "사회조사 결과 보도자료", "org": "통계청 · 2024. 11.", "type": "PDF",
     "desc": "전국 19,000 가구 36,000명 대상. 맞벌이 가구 자녀 돌봄 현황, 교육비 부담 등 학부모 인구통계 수치 근거."},
    {"name": "사회조사 통계표", "org": "통계청 · 2024. 11.", "type": "XLSX",
     "desc": "연령·성별·지역별 자녀 돌봄 이용 현황 원시 통계. 학부모 프로필 인구통계 수치 직접 근거."},
    {"name": "늘봄학교 운영 가이드라인", "org": "교육부 · 2024", "type": "PDF",
     "desc": "저녁늘봄 운영 시간·학급 편성 기준(20명 이내)·귀가 안전관리 절차·인력 운영 기준 등 정책 맥락 근거."},
]

@st.cache_data(show_spinner=False)
def load_docs():
    docs = {}
    for f in DATA_DIR.glob("*.pdf"):
        try:
            reader = PdfReader(str(f))
            docs[f.name] = "".join(p.extract_text() or "" for p in reader.pages)
        except: pass
    for f in DATA_DIR.glob("*.xlsx"):
        try:
            wb = openpyxl.load_workbook(str(f), data_only=True)
            text = ""
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    r = " | ".join(str(c) for c in row if c is not None)
                    if r.strip(): text += r + "\n"
            docs[f.name] = text
        except: pass
    return docs

def resolve_source_name(fname):
    """실제 파일명 → DATA_INFO 표시 이름 매칭.

    예전엔 '이름의 앞 두 단어 중 하나라도 파일명에 있으면 채택'이었는데,
    거의 모든 문서명이 "늘봄학교"로 시작해서 무조건 첫 번째 항목으로
    잘못 매칭됐다 (질적 사례연구든 가이드라인이든 다 "성과분석 연구"로
    표시되는 버그).

    지금은: 이름의 단어가 파일명과 "전부" 겹치는 항목을 최우선으로 찾고
    (완전 일치), 그런 항목이 여럿이면 매칭 글자 수 + 확장자(pdf/xlsx)
    일치 가산점이 가장 높은 걸 고른다. 완전 일치가 하나도 없으면
    부분 일치 중 최고 점수로 대체한다. (부분 점수만으로는 "교원 인식
    조사"처럼 세 단어가 다 맞는 문서가, 두 단어만 맞지만 우연히 글자
    수가 비슷한 다른 문서한테 동점으로 밀리는 경우가 있어서 완전 일치를
    먼저 본다.)
    """
    ext = Path(fname).suffix.lower().lstrip(".")
    full_matches = []
    partial_best_name, partial_best_score = None, 0
    for d in DATA_INFO:
        words = d["name"].split()
        matched = [w for w in words if w in fname]
        score = sum(len(w) for w in matched)
        if d.get("type", "").lower() == ext:
            score += 5
        if words and len(matched) == len(words):
            full_matches.append((score, d["name"]))
        if score > partial_best_score:
            partial_best_score = score
            partial_best_name = d["name"]
    if full_matches:
        full_matches.sort(reverse=True)
        return full_matches[0][1]
    return partial_best_name or fname[:20]

def search_docs(docs, query, n=3):
    kws = [w for w in query.split() if len(w) > 1]
    scored = [(sum(t.count(k) for k in kws), fn, t) for fn, t in docs.items()]
    scored = [(s,fn,t) for s,fn,t in scored if s > 0]
    scored.sort(reverse=True)
    results = []
    for _, fname, text in scored[:n]:
        chunks = []
        for kw in kws:
            idx = text.find(kw)
            while idx != -1 and len(chunks) < 2:
                chunk = text[max(0,idx-200):idx+400].strip()
                if chunk not in chunks: chunks.append(chunk)
                idx = text.find(kw, idx+1)
        if chunks:
            results.append({"source": resolve_source_name(fname), "text": "\n...\n".join(chunks[:2])})
    return results

def build_messages(persona_key, history, include_others):
    """세션 히스토리를 '이 페르소나'의 시점으로 재구성한다.

    include_others=False (기본 질문 라운드): 이 페르소나 자신의 과거 발언과
    진행자 질문만 보여준다. 다른 참여자의 답변은 아예 안 보이게 해서,
    질문 하나에 3명이 완전히 독립적으로 답하도록 한다.

    include_others=True ("반응 유도" 라운드에서만): 다른 참여자의 답변도
    화자 이름을 붙여 user 메시지(참고 정보)로 포함시켜서, 서로의 말에
    실제로 반응할 수 있게 한다.

    (참고: 예전엔 다른 참여자 답변을 이름 표시 없이 그냥 assistant로 넘겨서
    모델이 남의 말을 자기 말로 착각하는 버그가 있었음 — 지금은 이름을
    붙이고, 애초에 기본 라운드에서는 아예 안 보여주는 걸로 정리함.)
    """
    raw = []
    for m in history[-16:]:
        if m["role"] == "user":
            raw.append(("user", m["content"]))
        elif m.get("persona") == persona_key:
            raw.append(("assistant", m["content"]))
        elif include_others:
            speaker = PERSONA_BASE.get(m.get("persona"), {}).get("name", "다른 참여자")
            raw.append(("user", f"[{speaker}의 발언] {m['content']}"))
        # include_others=False면 다른 참여자 발언은 그냥 건너뜀

    merged = []
    for role, content in raw:
        if merged and merged[-1]["role"] == role:
            merged[-1]["content"] += "\n" + content
        else:
            merged.append({"role": role, "content": content})

    if not merged or merged[0]["role"] != "user":
        merged.insert(0, {"role": "user", "content": "(대화 시작)"})

    return merged


def get_response(persona_key, question, history, docs, mode, target, reaction=False):
    p = PERSONA_BASE[persona_key]
    if target != "전체" and p["name"] != target:
        return None

    ctx = search_docs(docs, question)
    ctx_text = "\n\n---\n\n".join(f"[출처: {c['source']}]\n{c['text'][:600]}" for c in ctx) if ctx else "이 질문에 대한 관련 데이터가 없습니다."
    sources = list({c["source"] for c in ctx})

    # 탐색/검증에 따라 다른 프로필 사용
    profile = p[f"{mode}_profile"]

    policy_sec = f"\n\n[정책 초안]\n{POLICY_DRAFT}\n" if mode == "검증" else ""
    others = ", ".join(f"{v['name']}({k})" for k,v in PERSONA_BASE.items() if k != persona_key)

    reaction_rule = (
        f"""5. 대화 중 "[이름의 발언] ..." 형태로 표시된 내용은 다른 참여자({others})가 방금 한 말입니다.
   해당 발언에 동의/반박하거나, 자신의 경험과 연결지어 자연스럽게 반응하세요.
   가능하면 단순 공감으로 끝내지 말고, 상대방에게 되묻거나 자신의 입장에서 구체적으로 덧붙이세요.
6. """
        if reaction else
        """5. 다른 참여자를 언급하거나 그들의 의견을 추측하지 말고, 오직 자신의 경험과 입장에서만 답하세요.
6. """
    )

    system = f"""당신은 늘봄학교 FGI에 참여하는 합성 사용자입니다.

{profile}
{policy_sec}
참고 데이터:
{ctx_text}

응답 규칙:
1. 위 프로필에 맞게 자연스럽게 대화하세요.
2. 참고 데이터에 있는 내용만 구체적 수치로 언급하세요.
3. 질문이 늘봄학교/돌봄 정책과 무관한 개인적인 내용(예: 식사, 취미, 가족, 연봉 등)이거나
   참고 데이터로 뒷받침할 수 없는 내용이면, 다른 말 절대 덧붙이지 말고 정확히 이 한 문장만 출력하세요:
   답변할 수 없습니다.
   금지 사항: "그건 답변드리기 어렵네요", "오늘 주제에 집중하죠", "(웃음)" 같은 식으로
   돌려서 거절하지 마세요. 반드시 "답변할 수 없습니다." 한 문장 그대로만 출력해야 합니다.
4. 탐색 모드에서 정책 초안 내용을 언급하는 질문을 받으면 반드시 "그런 계획이 있는지 몰랐어요"라고만 답하세요.
{reaction_rule}250자 내외로 간결하게 존댓말로 답변하세요."""

    messages = build_messages(persona_key, history, include_others=reaction)

    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY", "") or st.session_state.get("api_key", "")
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(model=MODEL, max_tokens=400, system=system, messages=messages)
        answer = resp.content[0].text.strip()
        # 거부 판정: "답변할 수 없습니다" 또는 "몰랐어요" 캔드 문구가 답변 "전체"에
        # 가까울 때만 거부로 본다. 정상 답변 안에 자연스럽게 그 표현이 섞여
        # 있는 경우(예: "미리 알고 계셨어요?"라는 질문에 "몰랐어요"로 답하며
        # 말을 이어가는 경우)까지 거부로 잘못 판정하지 않기 위함.
        stripped = answer.strip().rstrip(".。!? ")
        canned = "그런 계획이 있는지 몰랐어요"
        refused = (
            answer.strip() == "답변할 수 없습니다."
            or answer.strip() == "답변할 수 없습니다"
            or (stripped == canned and len(stripped) <= len(canned) + 3)
        )
        return {"answer": answer, "sources": sources, "refused": refused}
    except Exception as e:
        return {"answer": f"오류: {e}", "sources": [], "refused": True}

REACTION_PROMPTS = [
    "방금 다른 참여자들이 하신 이야기를 들으셨죠? 그 의견들에 대해 어떻게 생각하시는지 자유롭게 말씀해주세요.",
    "지금까지 나온 이야기 중에 동의하기 어려운 부분이나, 반대로 완전히 공감되는 부분이 있으면 말씀해주세요.",
    "다른 분들 의견 듣고 나서 생각이 좀 바뀌었거나, 덧붙이고 싶은 이야기가 있나요?",
]

def run_round(sess_key, question, docs, mode, target, shuffle=False, reaction=False):
    """질문 하나에 대해 참여자(들)가 순서대로 응답하는 한 라운드를 실행한다.
    - shuffle=True면 매번 응답 순서를 무작위로 섞어서, 항상 같은 사람이
      마지막에 '정리하는' 것처럼 보이는 걸 방지한다 (반응 유도 라운드용).
    - reaction=True로 저장된 메시지는 화면에서 "🗣️ 반응" 배지로 구분 표시된다.
    """
    st.session_state.sessions[sess_key].append({
        "role": "user", "content": question, "reaction_prompt": reaction
    })
    order = list(PERSONA_BASE.keys())
    if shuffle:
        random.shuffle(order)
    for pk in order:
        result = get_response(pk, question, st.session_state.sessions[sess_key], docs, mode, target, reaction=reaction)
        if result:
            st.session_state.sessions[sess_key].append({
                "role": "assistant", "content": result["answer"],
                "persona": pk, "sources": result["sources"],
                "refused": result["refused"], "reaction": reaction
            })
            time.sleep(0.1)

def make_log(pid, mode):
    msgs = st.session_state.sessions.get(f"{pid}_{mode}", [])
    lines = ["="*60, "늘봄학교 FGI 대화록",
             f"참여자: {pid} | 단계: {mode}",
             f"저장: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "="*60, ""]
    for m in msgs:
        if m["role"] == "user":
            tag = " (반응 유도)" if m.get("reaction_prompt") else ""
            lines.append(f"[진행자]{tag}\n{m['content']}\n")
        else:
            p = PERSONA_BASE.get(m.get("persona",""))
            nm = f"{p['emoji']} {p['name']}({m['persona']})" if p else "응답자"
            tag = " [반응]" if m.get("reaction") else ""
            lines.append(f"[{nm}]{tag}")
            lines.append(m["content"])
            if m.get("sources"): lines.append(f"출처: {', '.join(m['sources'])}")
            lines.append("")
    return "\n".join(lines)

def init():
    if "page" not in st.session_state: st.session_state.page = "fgi"
    if "mode" not in st.session_state: st.session_state.mode = "탐색"
    if "participant" not in st.session_state: st.session_state.participant = "A"
    if "sessions" not in st.session_state:
        st.session_state.sessions = {f"{p}_{m}": [] for p in PARTICIPANTS for m in MODES}
    if "target" not in st.session_state: st.session_state.target = "전체"
    if "start_times" not in st.session_state: st.session_state.start_times = {}
    if "api_key" not in st.session_state: st.session_state.api_key = ""

init()
docs = load_docs()

pid = st.session_state.participant
mode = st.session_state.mode
sess_key = f"{pid}_{mode}"
msgs = st.session_state.sessions[sess_key]
q_count = sum(1 for m in msgs if m["role"] == "user")
elapsed = int(time.time() - st.session_state.start_times.get(sess_key, time.time()))
elapsed_str = f"{elapsed//60:02d}:{elapsed%60:02d}"

# ══════════════════════════════════════
# FGI 화면
# ══════════════════════════════════════
if st.session_state.page == "fgi":

    pill_cls = "badge-blue" if mode == "탐색" else "badge-green"
    log_text = make_log(pid, mode)
    log_fname = f"FGI_{pid}_{mode}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    # 네비바
    nc1, nc2 = st.columns([3, 1])
    with nc1:
        st.markdown(f"""
        <div class="top-nav">
          <div style="display:flex;align-items:center">
            <div class="nav-logo">🏫 늘봄 FGI</div>
            <div class="nav-sep"></div>
            <div class="nav-ctx">참여자 {pid}</div>
            <div class="nav-sep"></div>
            <span class="badge {pill_cls}">{mode} 단계</span>
          </div>
        </div>""", unsafe_allow_html=True)
    with nc2:
        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button("⚙ 설정", key="btn_s", use_container_width=True):
                st.session_state.page = "settings"; st.rerun()
        with bc2:
            st.download_button("⬇ 저장", data=log_text, file_name=log_fname,
                               mime="text/plain", use_container_width=True, key="dl_m")

    sc, mc = st.columns([1, 4])

    with sc:
        # ── 모드 선택 (버튼 2개, 색상으로 선택 표시) ──
        st.markdown('<div class="sb-sec"><div class="sb-lbl">정책 단계</div>', unsafe_allow_html=True)
        m1, m2 = st.columns(2)
        with m1:
            탐색_style = "primary" if mode == "탐색" else "secondary"
            if st.button("탐색", key="mode_탐색", use_container_width=True,
                         type=탐색_style):
                st.session_state.mode = "탐색"; st.rerun()
        with m2:
            검증_style = "primary" if mode == "검증" else "secondary"
            if st.button("검증", key="mode_검증", use_container_width=True,
                         type=검증_style):
                st.session_state.mode = "검증"; st.rerun()

        # 선택 상태 텍스트
        if mode == "탐색":
            st.markdown('<div style="font-size:10px;color:#3b82f6;padding:2px 4px 8px">초안 비공개 상태</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div style="font-size:10px;color:#16a34a;padding:2px 4px 8px">초안 공개 상태</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── 이해관계자 ──
        st.markdown('<div class="sb-sec"><div class="sb-lbl">이해관계자</div>', unsafe_allow_html=True)
        for key, p in PERSONA_BASE.items():
            cnt = sum(1 for m in msgs if m.get("persona") == key)
            av = "av-a" if key=="학부모" else ("av-b" if key=="교사" else "av-c")
            st.markdown(f"""
            <div class="p-row">
              <div class="av {av}">{p['emoji']}</div>
              <div style="flex:1;min-width:0">
                <div class="p-nm">{p['name']}</div>
                <div class="p-rl">{key} · {p['region']}</div>
              </div>
              <div class="p-ct">{cnt}회</div>
            </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── 세션 ──
        st.markdown(f"""
        <div class="sb-sec">
          <div class="sb-lbl">세션</div>
          <div class="timer-row">
            <div class="timer">{elapsed_str}</div>
            <div class="qcnt">질문 {q_count}회</div>
          </div>
        </div>""", unsafe_allow_html=True)

        # ── 데이터 ──
        st.markdown('<div class="sb-sec"><div class="sb-lbl">데이터 (7)</div>', unsafe_allow_html=True)
        kb_html = ""
        for d in DATA_INFO:
            kb_html += f'<div class="kb-r"><div class="kb-dot"></div><div class="kb-nm">{d["name"]}</div></div>'
        st.markdown(kb_html + '</div>', unsafe_allow_html=True)

    with mc:
        # 검증 배너
        if mode == "검증":
            st.markdown("""
            <div class="policy-banner">
              <span style="font-size:16px">✅</span>
              <div>
                <div class="pb-title">저녁늘봄 전면 확대 시범사업(안) 적용 중</div>
                <div class="pb-sub">17시~20시 의무 운영 · 석식 무상 · 교사 업무 완전 제외</div>
              </div>
            </div>""", unsafe_allow_html=True)

        # 채팅
        if not msgs:
            st.info("💬 아래에 질문을 입력하면 합성 이해관계자 3인이 순차 응답합니다.")
        else:
            for m in msgs:
                if m["role"] == "user":
                    if m.get("reaction_prompt"):
                        # 반응 라운드는 '진행자가 새 질문을 던진 것'처럼 보이면
                        # 어색하니까, 대화 흐름 안의 작은 구분선으로만 표시.
                        # (모델에게 보내는 내부 지시문은 그대로 유지됨 — 화면 표시만 다름)
                        st.markdown(
                            '<div style="text-align:center;margin:14px 0;color:#9ca3af;'
                            'font-size:11px;display:flex;align-items:center;gap:8px">'
                            '<div style="flex:1;height:1px;background:#e5e7eb"></div>'
                            '🗣️ 참여자 반응'
                            '<div style="flex:1;height:1px;background:#e5e7eb"></div>'
                            '</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(
                            f'<div style="display:flex;justify-content:flex-end;margin:4px 0">'
                            f'<div class="q-bbl">{m["content"]}</div></div>',
                            unsafe_allow_html=True)
                else:
                    p = PERSONA_BASE.get(m.get("persona",""))
                    if not p: continue
                    av = "av-a" if m["persona"]=="학부모" else ("av-b" if m["persona"]=="교사" else "av-c")
                    react_badge = (
                        '<span style="font-size:10px;color:#7c3aed;background:#f3e8ff;'
                        'border-radius:8px;padding:1px 6px;margin-left:6px">🗣️ 반응</span>'
                    ) if m.get("reaction") else ""
                    if m.get("refused"):
                        body = ('<div class="a-refused">' + m["content"] + '</div>'
                                '<div class="refused-tag">⚠ 데이터 범위 외 — 답변 불가</div>')
                    else:
                        chips = "".join(
                            f'<span class="src-chip">📎 {s}</span>'
                            for s in m.get("sources",[])
                        )
                        src_row = f'<div class="src-row">{chips}</div>' if chips else ""
                        body = f'<div class="a-bbl">{m["content"]}</div>{src_row}'

                    st.markdown(f"""
                    <div class="a-blk" style="margin:6px 0">
                      <div class="a-hd">
                        <div class="av {av}" style="width:22px;height:22px;font-size:11px">{p['emoji']}</div>
                        <div class="a-nm">{p['name']} · {m['persona']}{react_badge}</div>
                      </div>
                      {body}
                    </div>""", unsafe_allow_html=True)

        st.markdown('<hr style="margin:8px 0;border-color:#e5e7eb">', unsafe_allow_html=True)

        # ── 대상 선택 (버튼, 선택된 것 파란색) ──
        targets = ["전체"] + [p["name"] for p in PERSONA_BASE.values()]
        tgt_cols = st.columns(len(targets))
        for i, t in enumerate(targets):
            with tgt_cols[i]:
                btn_type = "primary" if st.session_state.target == t else "secondary"
                if st.button(t, key=f"tgt_{t}", use_container_width=True, type=btn_type):
                    st.session_state.target = t; st.rerun()

        # ── 참여자 반응 유도 버튼 ──
        # 항상 보이지만, 대화가 없으면 비활성화 + 안내 문구.
        # 대상은 항상 "전체"로 강제(반응은 셋이 다 있어야 의미가 있음).
        # 순서를 매번 섞어서 항상 같은 사람이 마무리하는 것처럼 안 보이게 함.
        has_msgs = bool(msgs)
        if st.button("🗣️ 참여자 반응 유도", key="induce_reaction",
                     use_container_width=True, disabled=not has_msgs):
            if sess_key not in st.session_state.start_times:
                st.session_state.start_times[sess_key] = time.time()
            reaction_q = random.choice(REACTION_PROMPTS)
            with st.spinner("서로 반응하는 중..."):
                run_round(sess_key, reaction_q, docs, mode, "전체",
                          shuffle=True, reaction=True)
            st.rerun()
        if not has_msgs:
            st.caption("💡 질문을 먼저 하나 던지면, 참여자들이 서로 반응하게 만들 수 있어요.")

        question = st.chat_input("질문을 입력하세요...")
        if question:
            if sess_key not in st.session_state.start_times:
                st.session_state.start_times[sess_key] = time.time()
            with st.spinner("응답 생성 중..."):
                run_round(sess_key, question, docs, mode, st.session_state.target)
            st.rerun()

# ══════════════════════════════════════
# 설정 화면
# ══════════════════════════════════════
else:
    sc1, sc2 = st.columns([3,1])
    with sc1:
        st.markdown("""
        <div class="top-nav">
          <div class="nav-logo">🏫 늘봄 FGI
            <span style="font-size:12px;font-weight:400;color:#9ca3af;margin-left:4px">/ 설정</span>
          </div>
        </div>""", unsafe_allow_html=True)
    with sc2:
        if st.button("← FGI로 돌아가기", use_container_width=True):
            st.session_state.page = "fgi"; st.rerun()

    st.markdown('<div class="set-page">', unsafe_allow_html=True)

    # API 키
    if not (st.secrets.get("ANTHROPIC_API_KEY","") or st.session_state.api_key):
        with st.expander("🔑 API 키 설정"):
            k = st.text_input("Anthropic API 키", type="password", placeholder="sk-ant-...")
            if k: st.session_state.api_key = k

    # 참여자
    st.markdown('<div class="set-h">👥 참여자 관리</div>', unsafe_allow_html=True)
    st.markdown('<div class="set-sub">각 참여자 × 단계는 독립된 세션입니다. (A_탐색, A_검증, B_탐색 … 총 20개)</div>', unsafe_allow_html=True)

    ptab_html = '<div class="ptab-grid">'
    for p in PARTICIPANTS:
        on = "on" if p == pid else ""
        ptab_html += f'<span class="ptab {on}">{p}</span>'
    ptab_html += '</div>'
    st.markdown(ptab_html, unsafe_allow_html=True)

    pcols = st.columns(len(PARTICIPANTS))
    for i, p in enumerate(PARTICIPANTS):
        with pcols[i]:
            if st.button(p, key=f"ps_{p}", use_container_width=True):
                st.session_state.participant = p
                st.session_state.page = "fgi"; st.rerun()

    if st.button("🔄 현재 세션 초기화", key="reset_s"):
        st.session_state.sessions[sess_key] = []
        st.session_state.start_times.pop(sess_key, None)
        st.success(f"참여자 {pid} · {mode} 세션 초기화 완료")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # 합성 이해관계자
    st.markdown('<div class="set-h">🤖 합성 이해관계자</div>', unsafe_allow_html=True)
    st.markdown('<div class="set-sub">실제 공공데이터를 기반으로 설계된 3인의 이해관계자입니다. 유저 스토리 형식으로 작성되었습니다.</div>', unsafe_allow_html=True)

    for key, p in PERSONA_BASE.items():
        av = "av-a" if key=="학부모" else ("av-b" if key=="교사" else "av-c")
        story = PERSONA_STORIES[key].replace("\n", "<br>")
        tags_html = "".join(
            f'<span class="src-tag">{label} <span style="color:#9ca3af;font-size:11px">· {src}</span></span>'
            for label, src in p["tags_src"]
        )
        st.markdown(f"""
        <div class="pc-card">
          <div class="pc-top">
            <div class="pc-av {av}">{p['emoji']}</div>
            <div>
              <div class="pc-nm">{p['name']} · {key}</div>
              <div class="pc-mt">{p['age']}세 · {p['short']}</div>
            </div>
          </div>
          <div class="pc-body">{story}</div>
          <div class="tag-row">{tags_html}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # 정책 초안
    st.markdown('<div class="set-h">📄 정책 초안</div>', unsafe_allow_html=True)
    st.markdown('<div class="set-sub">검증 단계에서만 이해관계자에게 공개됩니다. 탐색 단계 이해관계자는 이 내용을 모릅니다.</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="policy-box">
      <div class="pol-title">저녁늘봄 전면 확대 시범사업(안) · 교육부 방과후돌봄지원과</div>
      <div class="pol-body">{POLICY_DRAFT.replace(chr(10), "<br>")}</div>
      <div class="pol-note">탐색 단계 이해관계자는 위 내용을 모르는 상태로 응답합니다.</div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # 데이터
    st.markdown('<div class="set-h">🗄 데이터 (7개)</div>', unsafe_allow_html=True)
    for d in DATA_INFO:
        bc = "pdf-b" if d["type"]=="PDF" else "xls-b"
        st.markdown(f"""
        <div class="kb-card">
          <div class="kbc-top">
            <div class="kbc-nm">{d['name']}</div>
            <div class="kbc-b {bc}">{d['type']}</div>
          </div>
          <div class="kbc-org">{d['org']}</div>
          <div class="kbc-desc">{d['desc']}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
