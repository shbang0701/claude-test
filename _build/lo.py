# -*- coding: utf-8 -*-
"""LibreOffice 실행 도우미 (다시 만들기 · 검수 전용 — 템플릿 사용에는 필요 없음).
PATH의 soffice를 임시 프로필로 실행한다.
소켓 제한이 있는 환경이면 LO_HELPER 환경 변수에 get_soffice_env()를 제공하는 soffice.py가 있는 폴더를 지정한다."""
import os
import subprocess
import sys
import tempfile
from pathlib import Path


def env():
    helper = os.environ.get("LO_HELPER")
    if helper:
        sys.path.insert(0, helper)
        from soffice import get_soffice_env
        return get_soffice_env()
    e = os.environ.copy()
    e.setdefault("SAL_USE_VCLPLUGIN", "svp")
    return e


def run(args, **kw):
    with tempfile.TemporaryDirectory(prefix="lo_profile_", ignore_cleanup_errors=True) as prof:
        return subprocess.run(["soffice", f"-env:UserInstallation={Path(prof).as_uri()}"] + list(args),
                              env=env(), **kw)


def convert(src, fmt, outdir):
    """src를 fmt(pdf, xlsx …)로 변환해 outdir에 저장하고 결과 경로를 돌려준다.
    xlsx → xlsx 변환은 값이 없는 수식을 계산해서 저장한다(재계산 사본 만들기)."""
    run(["--headless", "--convert-to", fmt, "--outdir", outdir, src], capture_output=True, timeout=300)
    return os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + "." + fmt.split(":")[0])
