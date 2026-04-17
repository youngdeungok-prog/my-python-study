import pandas as pd
import xlsxwriter
from datetime import datetime, timedelta

class GanttGenerator:
    """한국형 제조 스케줄링 간트차트 생성기"""
    
    def __init__(self, start_date_str='2026-04-17', days=15, shift_start='07:00', resolution=10, col_width=0.5):
        self.start_date_str = start_date_str
        self.days = days
        self.shift_start = shift_start
        self.resolution = resolution
        self.col_width = col_width
        self.base_start = datetime.strptime(f"{start_date_str} {shift_start}", "%Y-%m-%d %H:%M")
        self.slots_per_day = 144
        self.total_slots = days * self.slots_per_day
        self.grid_start_col = 4 # Factory, Op, Line, Seq
        
        # 색상 팔레트
        self.item_colors = ['#FFEB9C', '#C6EFCE', '#BDD7EE', '#FFCC99', '#CCFFCC', '#D9E1F2', '#E2EFDA']
        self.color_map = {}

    def _get_merge_ranges(self, target_df, col_name, parent_ranges=None):
        ranges = []
        total = len(target_df)
        if parent_ranges is None:
            start_idx = 0
            for i in range(1, total + 1):
                if i == total or target_df.loc[i, col_name] != target_df.loc[start_idx, col_name]:
                    ranges.append((start_idx, i - 1, target_df.loc[start_idx, col_name]))
                    start_idx = i
        else:
            for p_start, p_end, _ in parent_ranges:
                start_idx = p_start
                for i in range(p_start + 1, p_end + 2):
                    if i > p_end or target_df.loc[i, col_name] != target_df.loc[start_idx, col_name]:
                        ranges.append((start_idx, i - 1, target_df.loc[start_idx, col_name]))
                        start_idx = i
        return ranges

    def generate(self, line_master_df, plan_df, downtime_df=None, output_file='gantt_output.xlsx'):
        workbook = xlsxwriter.Workbook(output_file)
        ws = workbook.add_worksheet('GanttChart')
        
        # [서식 설정]
        formats = self._create_formats(workbook)
        
        # 1. 라인 마스터 확장 (Co-work 반영)
        expanded_master, _ = self._expand_master(line_master_df)
        
        # 2. 그리드 매트릭스 초기화 및 데이터 채우기
        matrix = [[None for _ in range(self.total_slots)] for _ in range(len(expanded_master))]
        line_to_row_idx = { (row['line_id'], row['seq']): i for i, row in expanded_master.iterrows() }
        
        self._fill_matrix(matrix, line_to_row_idx, plan_df, downtime_df, formats)
        
        # 3. 엑셀 출력: 헤더 및 시간 그리드
        self._write_headers(ws, formats)
        
        # 4. 엑셀 출력: 왼쪽 계층 레이블
        self._write_hierarchical_labels(ws, expanded_master, formats)
        
        # 5. 엑셀 출력: 매트릭스 데이터 (workbook 객체를 직접 전달하여 에러 해결!)
        self._write_gantt_bars(ws, workbook, matrix, expanded_master, formats)
        
        workbook.close()
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Gantt chart created: {output_file}")

    def _create_formats(self, wb):
        return {
            'header': wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BD', 'border': 1}),
            'date': wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#EBF1DE', 'border': 1}),
            'time': wb.add_format({'font_size': 7, 'valign': 'vcenter', 'top': 1, 'bottom': 1, 'left': 1, 'fg_color': '#F2F2F2'}),
            'hour_line': wb.add_format({'font_size': 7, 'valign': 'vcenter', 'top': 1, 'bottom': 1, 'left': 2, 'fg_color': '#F2F2F2'}),
            'cell': wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'pm': wb.add_format({'bg_color': '#D9D9D9', 'pattern': 4, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 8})
        }

    def _expand_master(self, df):
        expanded = []
        line_info = {}
        for _, row in df.iterrows():
            is_cw = row['co_work_yn'] == 'Y' and int(row['co_work_count']) > 1
            line_info[row['line_id']] = is_cw
            count = int(row['co_work_count']) if is_cw else 1
            for s in range(1, count + 1):
                expanded.append({'factory_id': row['factory_id'], 'op_id': row['op_id'], 'line_id': row['line_id'], 'seq': s, 'is_cowork': is_cw})
        return pd.DataFrame(expanded).sort_values(['factory_id', 'op_id', 'line_id', 'seq']).reset_index(drop=True), line_info

    def _fill_matrix(self, matrix, line_to_row_idx, plan_df, downtime_df, formats):
        # PM 채우기
        if downtime_df is not None:
            for _, dt in downtime_df.iterrows():
                for s in range(1, 10):
                    if (dt['line_id'], s) in line_to_row_idx:
                        r_idx = line_to_row_idx[(dt['line_id'], s)]
                        s_idx = int((dt['start_time'] - self.base_start).total_seconds() // (self.resolution * 60))
                        e_idx = int((dt['end_time'] - self.base_start).total_seconds() // (self.resolution * 60))
                        for c in range(max(0, s_idx), min(self.total_slots, e_idx)):
                            matrix[r_idx][c] = {'text': 'PM', 'fmt': formats['pm'], 'id': 'PM'}
        # 계획 채우기
        for _, p in plan_df.iterrows():
            key = (p['line_id'], p['seq'])
            if key in line_to_row_idx:
                r_idx = line_to_row_idx[key]
                item_id = p['item_id']
                if item_id not in self.color_map: self.color_map[item_id] = self.item_colors[len(self.color_map) % len(self.item_colors)]
                
                s_idx = int((p['start_time'] - self.base_start).total_seconds() // (self.resolution * 60))
                e_idx = int((p['end_time'] - self.base_start).total_seconds() // (self.resolution * 60))
                txt = f"{item_id}({p['qty']})\n{p['comment']}" if pd.notna(p['comment']) else f"{item_id}({p['qty']})"
                
                for c in range(max(0, s_idx), min(self.total_slots, e_idx)):
                    matrix[r_idx][c] = {'text': txt, 'color': self.color_map[item_id], 'id': p['mfg_order']}

    def _write_headers(self, ws, formats):
        labels = ['Factory', 'Operation', 'Line', 'Seq']
        for i, label in enumerate(labels):
            ws.merge_range(0, i, 1, i, label, formats['header'])
            ws.set_column(i, i, 4 if label == 'Seq' else 12)
        
        for d in range(self.days):
            d_col = self.grid_start_col + (d * self.slots_per_day)
            ws.merge_range(0, d_col, 0, d_col + self.slots_per_day - 1, (self.base_start + timedelta(days=d)).strftime("%Y-%m-%d (%a)"), formats['date'])
            for s in range(self.slots_per_day):
                c_idx = d_col + s
                ws.set_column(c_idx, c_idx, self.col_width)
                fmt = formats['hour_line'] if s % 6 == 0 else formats['time']
                ws.write(1, c_idx, self.shift_start if s == 0 else "", fmt)

    def _write_hierarchical_labels(self, ws, df, formats):
        f_ranges = self._get_merge_ranges(df, 'factory_id')
        op_ranges = self._get_merge_ranges(df, 'op_id', f_ranges)
        line_ranges = self._get_merge_ranges(df, 'line_id', op_ranges)
        for col_idx, ranges in enumerate([f_ranges, op_ranges, line_ranges]):
            for s_row, e_row, val in ranges:
                if s_row == e_row: ws.write(s_row + 2, col_idx, val, formats['cell'])
                else: ws.merge_range(s_row + 2, col_idx, e_row + 2, col_idx, val, formats['cell'])
        for i, row in df.iterrows():
            ws.write(i + 2, 3, row['seq'] if row['is_cowork'] else "", formats['cell'])

    def _write_gantt_bars(self, ws, workbook, matrix, df, formats):
        for r_idx in range(len(df)):
            excel_row = r_idx + 2
            ws.set_row(excel_row, 35)
            c_idx = 0
            while c_idx < self.total_slots:
                data = matrix[r_idx][c_idx]
                if data is None:
                    ws.write(excel_row, self.grid_start_col + c_idx, "", formats['cell'])
                    c_idx += 1
                else:
                    start_c = c_idx
                    curr_id = data['id']
                    while c_idx < self.total_slots and matrix[r_idx][c_idx] and matrix[r_idx][c_idx]['id'] == curr_id:
                        c_idx += 1
                    
                    if curr_id == 'PM':
                        use_fmt = formats['pm']
                    else:
                        # workbook 객체를 사용하여 동적으로 서식 생성
                        use_fmt = workbook.add_format({
                            'align': 'center', 'valign': 'vcenter', 'border': 1, 
                            'fg_color': data['color'], 'font_size': 8, 'bold': True, 'text_wrap': True
                        })
                    
                    col_pos = self.grid_start_col + start_c
                    if start_c == c_idx - 1:
                        ws.write(excel_row, col_pos, data['text'], use_fmt)
                    else:
                        ws.merge_range(excel_row, col_pos, excel_row, self.grid_start_col + c_idx - 1, data['text'], use_fmt)
                    
                    if curr_id != 'PM':
                        ws.write_comment(excel_row, col_pos, data['text'].strip())