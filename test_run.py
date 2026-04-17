from ko_schedule_gantt.core import GanttGenerator
import pandas as pd
from datetime import datetime, timedelta

# 1. 데이터 준비
line_m = pd.DataFrame({
    'factory_id': ['F1', 'F1'], 'op_id': ['OP10', 'OP20'], 'line_id': ['L101', 'L201'],
    'co_work_yn': ['Y', 'N'], 'co_work_count': [2, 1]
})

base = datetime(2026, 4, 17, 7, 0)
plans = pd.DataFrame({
    'line_id': ['L101', 'L101'], 'seq': [1, 2],
    'item_id': ['ITEM_A', 'ITEM_B'], 'mfg_order': ['MO01', 'MO02'],
    'qty': [100, 200], 'comment': ['정상', '긴급'],
    'start_time': [base + timedelta(hours=1), base + timedelta(hours=2)],
    'end_time': [base + timedelta(hours=6), base + timedelta(hours=8)]
})

# 2. 클래스 생성 및 실행
gen = GanttGenerator(start_date_str='2026-04-17', days=7)
gen.generate(line_m, plans, output_file='final_test.xlsx')