"""
설정 모듈
프로젝트 전역 경로 및 상수를 정의합니다.
PyInstaller frozen 실행 파일과 일반 실행 환경을 모두 지원합니다.
"""

import sys
import os

# --- 경로 설정 ---

def _get_base_dir() -> str:
    """프로젝트 루트 디렉토리를 반환합니다 (frozen exe 포함)."""
    if getattr(sys, 'frozen', False):
        # PyInstaller 번들 실행 시: 실행 파일이 있는 디렉토리
        return os.path.dirname(sys.executable)
    else:
        # 일반 실행 시: 이 파일의 상위 디렉토리 (프로젝트 루트)
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


BASE_DIR: str = _get_base_dir()

# 자동 저장 디렉토리
AUTOSAVE_DIR: str = os.path.join(BASE_DIR, "autosave")

# 규칙 파일 저장 디렉토리
RULES_DIR: str = os.path.join(AUTOSAVE_DIR, "rules")

# 검증 결과 저장 디렉토리
RESULTS_DIR: str = os.path.join(AUTOSAVE_DIR, "results")

# 이력 데이터 저장 디렉토리
HISTORY_DIR: str = os.path.join(AUTOSAVE_DIR, "history")

# React 프론트엔드 빌드 결과물 경로
FRONTEND_DIST: str = os.path.join(BASE_DIR, "frontend", "dist")

# --- 서버 설정 ---

# Flask 서버 포트
PORT: int = 5000

# Excel 데이터 변경 감지 폴링 간격 (초)
POLL_INTERVAL: int = 4

# --- 파일명 ---

# 기본 규칙 파일명
DEFAULT_RULES_FILENAME: str = "default_rules.json"

# 사용자 정의 규칙 파일명
CUSTOM_RULES_FILENAME: str = "custom_rules.json"

# 가격 이력 SQLite 데이터베이스 파일명
PRICE_HISTORY_DB: str = "price_history.db"
