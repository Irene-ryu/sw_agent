import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
from matplotlib.widgets import CheckButtons
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

class ExcelChartViewer:
    def __init__(self):
        self.df = None
        self.fig = None
        self.ax1 = None
        self.ax2 = None
        self.lines = {}
        self.line_visibility = {}
        
    def load_excel_file(self, file_path):
        """엑셀 파일 로드"""
        try:
            # 엑셀 파일에서 Sheet2 읽기
            self.df = pd.read_excel(file_path, sheet_name='Sheet2')
            print(f"파일 로드 성공: {file_path} (Sheet2)")
            print(f"컬럼: {list(self.df.columns)}")
            print(f"데이터 형태: {self.df.shape}")
            return True
        except Exception as e:
            print(f"파일 로드 오류: {e}")
            print("Sheet2를 찾을 수 없습니다. 다른 시트명을 확인해주세요.")
            return False
    
    def create_chart(self):
        """차트 생성"""
        if self.df is None:
            print("먼저 엑셀 파일을 로드해주세요.")
            return
        
        # 컬럼 인덱스 (0부터 시작)
        x_col = 0  # 1번째 컬럼 (x축)
        group_col1 = 1  # 2번째 컬럼
        group_col2 = 2  # 3번째 컬럼
        
        # y축 컬럼들 (4,5,6,7번째 컬럼)
        y_cols = [3, 4, 5, 6]  # 4,5,6,7번째 컬럼 인덱스
        
        # 기본 Y축 컬럼 찾기 ('파일 업로드 수' 또는 첫 번째 Y축 컬럼)
        default_y_col = None
        for y_col in y_cols:
            if y_col < len(self.df.columns):
                col_name = self.df.columns[y_col]
                if '파일 업로드 수' in col_name or 'upload' in col_name.lower():
                    default_y_col = y_col
                    break
        
        if default_y_col is None and y_cols[0] < len(self.df.columns):
            default_y_col = y_cols[0]  # 첫 번째 Y축 컬럼을 기본값으로
        
        # 컬럼명 가져오기
        x_col_name = self.df.columns[x_col]
        group_col1_name = self.df.columns[group_col1]
        group_col2_name = self.df.columns[group_col2]
        
        # 2,3번째 컬럼을 합쳐서 그룹 생성
        self.df['group'] = self.df[group_col1_name].astype(str) + ' - ' + self.df[group_col2_name].astype(str)
        
        # 그룹별로 데이터 분리
        groups = self.df['group'].unique()
        
        # 차트 생성 (단일축)
        self.fig, self.ax1 = plt.subplots(figsize=(12, 8))
        
        # 색상 팔레트 설정
        colors = plt.cm.tab10(np.linspace(0, 1, len(groups) * len(y_cols)))
        color_idx = 0
        
        # 마커 스타일
        markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p']
        
        # 모든 Y축 컬럼에 대해 라인 생성 (기본적으로 숨김)
        for i, group in enumerate(groups):
            group_data = self.df[self.df['group'] == group]
            
            if len(group_data) > 0:
                for j, y_col in enumerate(y_cols):
                    if y_col < len(self.df.columns):
                        y_col_name = self.df.columns[y_col]
                        # marker = markers[j % len(markers)]
                        
                        # 기본 Y축 컬럼만 표시, 나머지는 숨김
                        visible = (y_col == default_y_col)
                        
                        line, = self.ax1.plot(group_data[x_col_name], group_data[y_col_name], 
                                             linewidth=2,                       # marker=marker, 
                                             label=f'{group} ({y_col_name})',
                                             color=colors[color_idx],
                                             visible=visible)
                        self.lines[f'{group}_{y_col_name}'] = line
                        self.line_visibility[f'{group}_{y_col_name}'] = visible
                        color_idx += 1
        
        # 축 레이블 설정
        self.ax1.set_xlabel(x_col_name, fontsize=12)
        if default_y_col is not None:
            self.ax1.set_ylabel(self.df.columns[default_y_col], fontsize=12)
        
        # Y축 컬럼명들을 제목에 표시
        y_col_names = []
        for y_col in y_cols:
            if y_col < len(self.df.columns):
                y_col_names.append(self.df.columns[y_col])
        
        default_y_name = self.df.columns[default_y_col] if default_y_col is not None else "기본값"
        self.ax1.set_title(f'엑셀 데이터 라인 차트 - 기본 Y축: {default_y_name} (체크박스로 Y축 변경 가능)', fontsize=14, fontweight='bold')
        
        # 그리드 추가
        self.ax1.grid(True, alpha=0.3)
        
        # 초기 Y축 범위 설정
        self.update_y_axis_range()
        
        # 범례 생성
        self.create_legend_checkboxes()
        
        plt.tight_layout()
        plt.show()
    
    def create_legend_checkboxes(self):
        """범례 필터링을 위한 체크박스 생성"""
        if not self.lines:
            return
        
        # Y축 선택을 위한 체크박스 (좌측 하단)
        ax_y_checkbox = plt.axes([0.02, 0.02, 0.25, 0.15])
        
        # Y축 컬럼들만 추출
        y_cols = [3, 4, 5, 6]  # 4,5,6,7번째 컬럼 인덱스
        y_col_names = []
        y_col_visibility = []
        
        for y_col in y_cols:
            if y_col < len(self.df.columns):
                y_col_name = self.df.columns[y_col]
                y_col_names.append(y_col_name)
                # 기본 Y축 컬럼만 체크
                default_visible = any(f'_{y_col_name}' in key and self.line_visibility[key] for key in self.lines.keys())
                y_col_visibility.append(default_visible)
        
        # Y축 선택 체크박스 생성
        y_check = CheckButtons(ax_y_checkbox, y_col_names, y_col_visibility)
        
        def y_checkbox_callback(label):
            """Y축 체크박스 클릭 시 호출되는 함수"""
            # 해당 Y축 컬럼의 모든 라인을 토글
            for line_key, line in self.lines.items():
                if f'_{label}' in line_key:
                    if line.get_visible():
                        line.set_visible(False)
                        self.line_visibility[line_key] = False
                    else:
                        line.set_visible(True)
                        self.line_visibility[line_key] = True
            
            # Y축 범위 동적 조정
            self.update_y_axis_range()
            plt.draw()
        
        y_check.on_clicked(y_checkbox_callback)
        
        # 그룹별 라인 선택을 위한 체크박스 (좌측 중간)
        ax_group_checkbox = plt.axes([0.02, 0.20, 0.25, 0.25])
        
        # 그룹별 라인만 추출
        group_lines = {}
        for line_key, line in self.lines.items():
            # Y축 컬럼명 제거하여 그룹명만 추출
            group_name = line_key.split('_')[0]
            if group_name not in group_lines:
                group_lines[group_name] = []
            group_lines[group_name].append(line_key)
        
        group_labels = list(group_lines.keys())
        group_visibility = [True] * len(group_labels)  # 모든 그룹 기본 표시
        
        # 그룹별 체크박스 생성
        group_check = CheckButtons(ax_group_checkbox, group_labels, group_visibility)
        
        def group_checkbox_callback(label):
            """그룹 체크박스 클릭 시 호출되는 함수"""
            # 해당 그룹의 모든 라인을 토글
            for line_key in group_lines[label]:
                if line_key in self.lines:
                    line = self.lines[line_key]
                    if line.get_visible():
                        line.set_visible(False)
                        self.line_visibility[line_key] = False
                    else:
                        line.set_visible(True)
                        self.line_visibility[line_key] = True
            
            # Y축 범위 동적 조정
            self.update_y_axis_range()
            plt.draw()
        
        group_check.on_clicked(group_checkbox_callback)
    
    def update_y_axis_range(self):
        """Y축 범위를 동적으로 업데이트"""
        if not self.lines:
            return
        
        # 현재 표시된 라인들의 데이터 수집
        visible_data = []
        
        for line_key, line in self.lines.items():
            if line.get_visible():
                # 라인의 데이터 가져오기
                y_data = line.get_ydata()
                if len(y_data) > 0:
                    visible_data.extend(y_data)
        
        if visible_data:
            # 표시된 데이터의 최소/최대값 계산
            min_val = min(visible_data)
            max_val = max(visible_data)
            
            # 여백 추가 (10%)
            margin = (max_val - min_val) * 0.1
            if margin == 0:  # 모든 값이 같을 경우
                margin = abs(min_val) * 0.1 if min_val != 0 else 1
            
            y_min = min_val - margin
            y_max = max_val + margin
            
            # Y축 범위 설정
            self.ax1.set_ylim(y_min, y_max)
            
            # Y축 레이블 업데이트
            visible_y_cols = set()
            for line_key, line in self.lines.items():
                if line.get_visible():
                    # Y축 컬럼명 추출
                    y_col_name = line_key.split('_', 1)[1] if '_' in line_key else line_key
                    visible_y_cols.add(y_col_name)
            
            if len(visible_y_cols) == 1:
                # 하나의 Y축만 표시된 경우
                self.ax1.set_ylabel(list(visible_y_cols)[0], fontsize=12)
            else:
                # 여러 Y축이 표시된 경우
                self.ax1.set_ylabel('값', fontsize=12)
    
    def save_chart(self, file_path):
        """차트를 이미지로 저장"""
        if self.fig:
            self.fig.savefig(file_path, dpi=300, bbox_inches='tight')
            print(f"차트가 저장되었습니다: {file_path}")
    


def select_file():
    """파일 선택 다이얼로그"""
    root = tk.Tk()
    root.withdraw()  # 메인 창 숨기기
    
    file_path = filedialog.askopenfilename(
        title="엑셀 파일을 선택하세요",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    
    return file_path

def main():
    """메인 함수"""
    print("=== 엑셀 파일 라인 차트 생성기 ===")
    
    # 파일 선택
    file_path = select_file()
    
    if not file_path:
        print("파일이 선택되지 않았습니다.")
        return
    
    # 차트 뷰어 생성
    viewer = ExcelChartViewer()
    
    # 파일 로드
    if viewer.load_excel_file(file_path):
        # 차트 생성 (4,5,6,7번째 컬럼 모두 사용)
        viewer.create_chart()
        
        # 차트 저장 여부 확인
        save_choice = input("\n차트를 저장하시겠습니까? (y/n): ").lower()
        if save_choice == 'y':
            save_path = file_path.replace('.xlsx', '_chart.png').replace('.xls', '_chart.png')
            viewer.save_chart(save_path)
    else:
        print("파일 로드에 실패했습니다.")

if __name__ == "__main__":
    main()
