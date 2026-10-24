# scripts

복사 한 번으로 끝나는 PowerShell 7 스크립트입니다. 규칙은 [AGENTS.md](../AGENTS.md)에 있습니다.

## 분류

스크립트가 생기면 그때 폴더를 만듭니다.

- `office/` — Excel·Word·PPT 만들기·채우기 (COM)
- `files/` — 파일·폴더 정리, 이름 바꾸기, 비교
- `system/` — PC·환경 확인 (디스크, 설치 프로그램, 설정)
- `network/` — 연결·포트 확인
- `text/` — 로그·텍스트 가공

## 새 스크립트 만들기

1. [_template.ps1](_template.ps1)을 복사해서 시작합니다.
2. 이름은 `동사-명사.ps1`로 짓습니다 (예: `Export-SheetList.ps1`, `Get-DiskUsage.ps1`).
   동사는 PowerShell 표준 동사를 씁니다: Get, Set, New, Export, Import, Convert, Find, Test 등.
3. 아래 **목록**과 루트 [README.md](../README.md)에 한 줄씩 추가합니다.

## 업무 PC에서 실행하기

1. GitHub에서 파일 전체를 복사해, 업무 PC에 같은 이름으로 저장합니다 (UTF-8).
2. 맨 위 **설정** 블록만 고칩니다.
3. `pwsh`에서 `.\스크립트이름.ps1`로 실행합니다.
   - 도움말 보기: `Get-Help .\스크립트이름.ps1 -Examples`
   - 저장하지 않고 할 일만 보기: `.\스크립트이름.ps1 -WhatIf`

## 목록

- [network/Watch-MultiPing.ps1](network/Watch-MultiPing.ps1) — 여러 서버에 ping을 계속 보내 응답 · 끊김을 한 화면에 보여 주고, 상태가 바뀐 시각을 CSV로 남깁니다. ping이 막힌 서버는 `주소:포트`로 적어 TCP 연결로 확인합니다.
  - 대상 목록은 스크립트 옆 `ping-hosts.txt`에 한 줄에 하나씩 `이름 주소`로 적습니다. Excel에서 이름 · 주소 두 열을 복사해 붙여 넣어도 됩니다.

    ```text
    # 점검 대상 (# 뒤는 메모)
    SRV-01       10.0.0.10
    SRV-01-iLO   10.0.1.10
    SRV-02       10.0.0.11:22    # ping이 막혀 SSH 포트로 확인
    ```
  - 실행 기록은 `ping-logs\ping_<날짜_시각>.csv`, 끝내면 `..._요약.csv`가 생깁니다. Q 또는 Esc로 끝냅니다.
  - 실제 목록 · 기록은 저장소에 올리지 않습니다(`.gitignore`).
