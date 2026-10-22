# 명령어 모음

PowerShell 7(`pwsh`) 기준이다. 코드 블록 오른쪽 위 복사 버튼으로 복사해서 VS Code 터미널에 붙여넣는다.

**모든 명령은 WorkTool 폴더에서 실행한다.** VS Code로 WorkTool 폴더를 열면 터미널이 그 폴더에서 시작한다. 다른 곳이면 먼저 `cd`로 이동한다.

## 설치

### 폴더 만들기 (처음 한 번)

```powershell
New-Item -ItemType Directory -Force -Path modules, data | Out-Null
```

### 파일 붙여넣기

1. VS Code에서 파일 > 폴더 열기 → WorkTool 폴더
2. 탐색기에서 새 파일 → GitHub와 같은 경로·이름 (예: `Start-WorkTool.ps1`, `modules\Storage.psm1`)
3. GitHub 파일 화면의 "Copy raw file" 버튼으로 복사 → 붙여넣기 → 저장

## 실행

F5(디버그 실행)는 쓰지 않는다.

```powershell
.\Start-WorkTool.ps1
```

창이 떠 있는 동안 터미널은 프롬프트가 안 돌아온다 (정상). 창의 X는 트레이로 숨기기라서, 완전히 끄려면 **트레이 아이콘 우클릭 → 종료**. 그러면 프롬프트가 돌아온다.

## 단계별 테스트

### Storage

`modules\Storage.psm1`을 붙여넣은 뒤 한 블록씩 실행한다.

불러오고 테스트 기록 1건 저장:

```powershell
Import-Module .\modules\Storage.psm1 -Force
Save-Records test @([pscustomobject]@{ Id = (New-RecordId); Title = '테스트'; Body = "첫 줄`n둘째 줄, 쉼표" })
Get-Records test | Format-List
```

한 건 더 붙여서 저장 → 파일 2개(`test.csv`, `test.csv.bak`)와 건수 2가 나오면 정상:

```powershell
$r = @(Get-Records test)
Save-Records test @($r + [pscustomobject]@{ Id = (New-RecordId); Title = '두번째'; Body = 'x' })
Get-ChildItem .\data | Select-Object Name
@(Get-Records test).Count
```

테스트 파일 지우기:

```powershell
Remove-Item .\data\test.csv*
```

### 메모

붙여넣을 파일: `modules\Memo.psm1`, `modules\Storage.psm1`, `Start-WorkTool.ps1`. 그다음 실행한다.

```powershell
.\Start-WorkTool.ps1
```

확인할 것:

1. 왼쪽 메뉴에 "메모"가 있다
2. 제목·본문 쓰고 Ctrl+S → 왼쪽 목록에 뜨고 "저장됨 시각"이 나온다
3. 새 메모 → 하나 더 쓰고 목록에서 첫 메모 클릭 → 두 번째 메모가 자동 저장된다
4. 제목을 비우고 저장 → 저장 시각이 제목이 된다
5. 검색칸에 글자 입력 → 목록이 걸러진다
6. 아무 메모나 고치고 저장 안 한 채 창 닫기 → 다시 열면 고친 내용이 남아 있다
7. 삭제 → 확인 창 → 목록에서 사라진다
8. `data\memo.csv`를 Excel로 열어둔 채 저장 → 경고 창이 뜬다 (Excel 닫고 다시 저장하면 된다)

### 왼쪽 메뉴·테마

붙여넣을 파일: `modules\UiHelper.psm1`(새 파일), `Start-WorkTool.ps1`(수정), `modules\Memo.psm1`(수정).

확인할 것:

1. 왼쪽에 하늘색 메뉴(WorkTool, 날짜, 홈, 메모)가 있고 아이콘이 보인다
2. 메뉴를 누르면 화면이 바뀌고, 고른 메뉴는 진하게 + 왼쪽에 파란 막대
3. 최대화 버튼이 되고, F11로 전체 화면 켜고 끄기가 된다
4. 아래 상태 표시줄에 지금 쓰는 글꼴 이름이 나온다
5. 메모 기능은 전과 똑같이 된다

PC에 설치된 한글 글꼴 중 쓸 만한 것 확인 (툴이 고를 수 있는 후보):

```powershell
Add-Type -AssemblyName System.Drawing
[System.Drawing.Text.InstalledFontCollection]::new().Families.Name | Where-Object { $_ -match 'Pretendard|SUIT|Noto|Nanum|나눔|Spoqa|맑은|Malgun|Gothic|고딕' }
```

### 근태·메뉴 접기

붙여넣을 파일: `modules\Attendance.psm1`(새 파일), `modules\UiHelper.psm1`(수정), `Start-WorkTool.ps1`(수정).

확인할 것:

1. 메뉴 맨 위 `≡ WorkTool` 클릭 또는 Ctrl+B → 메뉴가 아이콘만 남게 접히고, 다시 누르면 펼쳐진다
2. 접은 상태에서 아이콘에 마우스를 올리면 메뉴 이름이 뜬다
3. 메뉴에 "근태"가 있고 들어가면 이번 달이 보인다
4. "지금 출근" → 목록에 오늘 출근이 생긴다. 한 번 더 누르면 "이미 있습니다"
5. 종류를 바꾸면 칸이 바뀐다 (출장은 종료일·장소, 업무외 근무는 끝 시각, 휴가는 종류 선택)
6. 업무외 근무 22:00~02:30 저장 → "(다음날) · 4:30", 위 요약 시간에 더해진다
7. ◀ ▶로 달 이동, "이번 달"로 돌아오기
8. 목록에서 고르면 오른쪽에 뜨고, 고쳐서 저장 / 삭제가 된다

### 일정·알람·트레이

붙여넣을 파일: `modules\Schedule.psm1`(새 파일), `modules\UiHelper.psm1`(수정), `Start-WorkTool.ps1`(수정).

확인할 것:

1. 창·작업 표시줄·트레이에 파란 W 아이콘이 보인다
2. 메뉴에 "일정"이 있고, 달력에서 날짜를 누르면 오른쪽 날짜가 바뀐다
3. 지금부터 2~3분 뒤 시각으로 알람 "정시" 일정 저장 → 달력에 그날이 굵게, 왼쪽 "다가오는 7일"에 보인다
4. 창의 X → 트레이로 숨겨지고 안내 풍선이 뜬다
5. 그 시각이 되면 (15초 안쪽 오차) 소리와 함께 오른쪽 아래 알림 창이 뜬다. "5분 뒤 다시" → 5분 뒤 다시 뜬다
6. 트레이 아이콘 더블클릭 → 창이 다시 뜬다
7. 반복 "평일"로 하나 만들고 토요일·월요일을 눌러 보기 → 월요일에만 보인다
8. 트레이 우클릭 → 종료 → 아이콘이 사라지고 터미널 프롬프트가 돌아온다
9. 다시 실행해도 이미 울린 알람은 또 울리지 않는다
