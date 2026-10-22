@AGENTS.md

## Claude 전용

- `docs/templates/infra-kit/`의 Excel · PPT · 아이콘 · PDF를 만들거나 고칠 때는 먼저 `infra-design` 스킬(`.claude/skills/infra-design/`)을 읽고 그 기준을 따른다.
- WorkTool 화면의 색 · 글꼴 · 크기는 `apps/WorkTool/modules/UiHelper.psm1`에서 정한다. `infra-design` 팔레트를 섞지 않는다.
