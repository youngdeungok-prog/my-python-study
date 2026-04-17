# ko-schedule-gantt-chart 📊

**한국형 제조/SCM 실무에 최적화된 엑셀 간트차트 생성 라이브러리** A specialized Excel Gantt chart generator tailored for Korean manufacturing SCM scheduling.

## ✨ Key Features (주요 기능)

- **Co-work & Multi-Sequence Support**: 동일 라인 내 여러 작업대(Sequence)에서의 동시 생산 공정을 완벽하게 시각화합니다.
- **10-Minute High Resolution**: 10분 단위의 세밀한 그리드를 제공하며, 가독성을 위해 1시간 단위로 굵은 가이드라인을 표시합니다.
- **Shift-Aware Scheduling**: 공장별 조업 시작 시각(예: 07:00)을 기준으로 하루의 타임라인을 재설정합니다.
- **Smart Tooltips (Comments)**: 셀 폭이 좁아 내용이 잘리는 짧은 공정도 마우스 오버 시 전체 정보(Item, Qty, Comment)를 확인할 수 있습니다.
- **Layered Rendering**: 비가동(PM/휴무) 구간을 먼저 렌더링하고 그 위에 생산 계획을 덮어쓰는 레이어 방식을 채택하여 데이터 충돌을 방지합니다.

## 📂 Project Structure

- `ko_schedule_gantt/`: 라이브러리 핵심 패키지
  - `core.py`: 간트차트 생성 로직이 담긴 `GanttGenerator` 클래스
- `test_run.py`: 라이브러리 구동 및 결과 확인을 위한 샘플 스케립트

## 📊 Data Schema (데이터 규격)

### 1. Line Master (라인 마스터)
왼쪽 레이블 영역의 구조를 정의합니다.
- `factory_id`, `op_id`, `line_id`: 계층 구조 정보
- `co_work_yn`: 동시 생산 여부 (Y/N)
- `co_work_count`: 동시 생산 가능 수 (Seq 생성 기준)

### 2. Plan Data (생산 계획)
차트 본문에 그려질 데이터입니다.
- `line_id`, `seq`: 타겟 위치 정보
- `item_id`, `qty`, `comment`: 바(Bar) 내부에 표기될 정보
- `start_time`, `end_time`: 작업 시간 (Datetime 형식)

## 🚀 Quick Start (사용법)

```python
from ko_schedule_gantt.core import GanttGenerator

# 생성기 초기화 (시작일, 기간, 조업시작시간 등 설정)
gen = GanttGenerator(start_date_str='2026-04-17', days=7, shift_start='07:00')

# 간트차트 생성
gen.generate(line_master_df, plan_df, downtime_df=pm_df, output_file='My_Gantt_Chart.xlsx')