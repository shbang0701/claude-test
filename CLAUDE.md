@AGENTS.md

## Claude 전용

- `docs/templates/infra-kit/`의 Excel · PPT · 아이콘 · PDF를 만들거나 고칠 때는 먼저 `infra-design` 스킬(`.claude/skills/infra-design/`)을 읽고 그 기준을 따른다.
- WorkTool 화면의 색 · 글꼴 · 크기 기준은 `apps/WorkTool/design/README.md`(시안)다. 지금 WinForms 코드(`apps/WorkTool/modules/UiHelper.psm1`)는 이전 하늘색 값을 쓰고 있고, WPF로 옮기면서 이 기준으로 바꾼다. `infra-design` 팔레트를 섞지 않는다.
