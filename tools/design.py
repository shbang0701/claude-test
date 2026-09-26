# -*- coding: utf-8 -*-
"""IT 인프라 문서 라이브러리 - 공통 디자인 토큰.
Word / PowerPoint / Excel 세 형식이 같은 색·글꼴·간격 체계를 쓰도록 한 곳에서 관리한다."""

# ── 색 (모두 RGB Hex, '#' 없음) ──────────────────────────────────────────────
INK        = "1B1F24"   # 본문 글자
INK_SOFT   = "5A6570"   # 보조 설명, 캡션
INK_FAINT  = "8A94A0"   # 라벨, 비활성
LINE       = "D5DBE1"   # 표 안쪽 선
LINE_SOFT  = "E7EBEF"   # 아주 옅은 구분선
SURFACE    = "F5F7F9"   # 연한 바탕
WHITE      = "FFFFFF"

PRIMARY      = "14466B" # 주색 (제목, 표 머리)
PRIMARY_MID  = "2C6E9B" # 보조 주색 (선, 강조)
PRIMARY_TINT = "E9F0F6" # 주색 바탕
PRIMARY_DEEP = "0E3350"

CODE_BG   = "F1F5F8"; CODE_TXT  = "17324A"; CODE_BDR = "DCE4EB"
OUT_BG    = "FAFBFC"; OUT_TXT   = "4A545E"; OUT_BDR  = "E4E8EC"

WARN      = "9A6207"; WARN_BG   = "FCF4E6"; WARN_BDR = "E8CFA0"   # 주의
DANGER    = "A32A30"; DANGER_BG = "FBEDED"; DANGER_BDR = "E9C3C4"  # 경고
OK        = "1B6E4B"; OK_BG     = "EAF4EF"; OK_BDR   = "BEDCCC"   # 정상/확인
NOTE      = "2C6E9B"; NOTE_BG   = "EFF4F8"; NOTE_BDR = "CFDCE7"   # 참고
VAR       = "8A5000"; VAR_BG    = "FFF3E0"                         # 치환값

# 상태 색 (PPT/Excel 공용)
ST_OK = "2E8B57"; ST_WARN = "D19B10"; ST_BAD = "C0392B"; ST_IDLE = "9AA5B1"

# ── 글꼴 ────────────────────────────────────────────────────────────────────
FONT_LATIN = "Malgun Gothic"
FONT_EA    = "맑은 고딕"
FONT_MONO  = "Consolas"
FONT_MONO_EA = "맑은 고딕"

# ── 단위 변환 ───────────────────────────────────────────────────────────────
def cm(v):   return int(round(v * 567))      # cm  -> twips
def pt(v):   return int(round(v * 20))       # pt   -> twips (간격)
def hp(v):   return int(round(v * 2))        # pt   -> half-point (글자 크기)
def eighth(v): return int(round(v * 8))      # pt   -> 1/8 pt (테두리 두께)
