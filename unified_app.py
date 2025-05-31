import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox,
    QScrollArea, QSizePolicy, QMainWindow, QGroupBox, QTabWidget, QDialog,
    QHeaderView, QTextEdit, QListWidget, QListWidgetItem, QSpinBox,
    QDoubleSpinBox, QComboBox
)
from PyQt6.QtCore import Qt, QLocale
from PyQt6.QtGui import QDoubleValidator, QFont, QTextCursor

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei'] 
plt.rcParams['axes.unicode_minus'] = False  

from unified_calculator import calculate_unified_efficiencies, ExperimentConfig, BatchExperimentAnalyzer, calculate_simple_efficiency
from factor_calculator import calculate_single_efficiency, calculate_factor_experiment
import numpy as np
import os

class MatplotlibCanvas(FigureCanvas):
 
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        FigureCanvas.setSizePolicy(self, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        FigureCanvas.updateGeometry(self)

    def plot(self, x_data, y_data, title="", x_label="", y_label="", legend_label="", color=None):
        self.axes.cla() 
        if x_data is not None and y_data is not None and len(x_data) > 0 and len(y_data) > 0:
            self.axes.plot(x_data, y_data, label=legend_label, color=color)
            if legend_label:
                self.axes.legend()
        self.axes.set_title(title)
        self.axes.set_xlabel(x_label)
        self.axes.set_ylabel(y_label)
        self.axes.grid(True)
        self.draw()

    def plot_comparison(self, datasets, title="", x_label="", y_label=""):
    
        self.axes.cla()
        colors = ['blue', 'red', 'green', 'orange']
        for i, (x_data, y_data, label) in enumerate(datasets):
            if x_data is not None and y_data is not None and len(x_data) > 0:
                self.axes.plot(x_data, y_data, label=label, color=colors[i % len(colors)], alpha=0.7)
        self.axes.set_title(title)
        self.axes.set_xlabel(x_label)
        self.axes.set_ylabel(y_label)
        self.axes.legend()
        self.axes.grid(True)
        self.draw()

class UnifiedMotorAnalysisApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("电机效率统一分析系统 - 综合实验平台")
        self.setGeometry(50, 50, 1500, 950)
        self.zheng_file = None
        self.fan_file = None
        self.batch_config = None 
        self.batch_analyzer = None  
        self.batch_file_list = [] 
        self.default_params = {
            "reference_v": "0.185",
            "initial_v": "2.52", 
            "r_load": "3.500",
            "drive_v": "12.000",
            "power_input": "13.0",
            "sampling_freq": "87500"
        }

        self.results = None
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(100)
        
        QLocale.setDefault(QLocale(QLocale.Language.C, QLocale.Country.AnyCountry))
        self._init_ui()

    def log(self, message, level="INFO"):
        """添加日志信息"""
        color_map = {
            "INFO": "black",
            "SUCCESS": "green",
            "WARNING": "orange",
            "ERROR": "red"
        }
        color = color_map.get(level, "black")
        self.log_text.append(f'<span style="color: {color};">[{level}] {message}</span>')

        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_text.setTextCursor(cursor)

    def _init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        title_label = QLabel("🔬 电机效率统一分析系统 - 综合实验平台")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        self.main_tabs = QTabWidget()
        dual_motor_widget = self._create_dual_motor_widget()
        self.main_tabs.addTab(dual_motor_widget, "🔄 双机标定实验")
        batch_experiment_widget = self._create_batch_experiment_widget()
        self.main_tabs.addTab(batch_experiment_widget, "🔍 批量实验分析")
        main_layout.addWidget(self.main_tabs)
        log_group = QGroupBox("系统日志")
        log_layout = QVBoxLayout()
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        self.log("系统已就绪，请选择实验类型开始分析", "INFO")
    def _create_dual_motor_widget(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        left_panel = self._create_control_panel()
        layout.addWidget(left_panel)
        right_panel = self._create_results_panel()
        layout.addWidget(right_panel, 1)
        return widget
    def _create_batch_experiment_widget(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        left_panel = self._create_batch_control_panel()
        layout.addWidget(left_panel)
        right_panel = self._create_batch_results_panel()
        layout.addWidget(right_panel, 1)
        return widget
    def _create_batch_control_panel(self):
        """创建批量实验控制面板"""
        panel = QWidget()
        panel.setFixedWidth(450)
        layout = QVBoxLayout(panel)
        explore_group = QGroupBox("🎯 探究类型选择")
        explore_layout = QVBoxLayout()
        self.batch_explore_type = QComboBox()
        self.batch_explore_type.addItems(["输入电压影响", "负载电阻影响", "磁场距离影响"])
        self.batch_explore_type.currentTextChanged.connect(self._on_batch_explore_type_changed)
        explore_layout.addWidget(QLabel("选择探究因素："))
        explore_layout.addWidget(self.batch_explore_type)
        self.fixed_params_label = QLabel("固定参数：负载电阻 = 10.0 Ω")
        self.fixed_params_label.setStyleSheet("color: blue; padding: 5px;")
        explore_layout.addWidget(self.fixed_params_label)
        explore_group.setLayout(explore_layout)
        layout.addWidget(explore_group)
        params_group = QGroupBox("📋 实验参数配置")
        params_layout = QVBoxLayout()
        self.batch_params_table = QTableWidget()
        self.batch_params_table.setColumnCount(3)
        self.batch_params_table.setHorizontalHeaderLabels(["序号", "变量值", "输入功率(W)"])
        self.batch_params_table.setMaximumHeight(200)
        param_btn_layout = QHBoxLayout()
        self.btn_add_batch_param = QPushButton("➕ 添加参数组")
        self.btn_add_batch_param.clicked.connect(self._add_batch_param)
        self.btn_remove_batch_param = QPushButton("➖ 删除选中组")
        self.btn_remove_batch_param.clicked.connect(self._remove_batch_param)
        param_btn_layout.addWidget(self.btn_add_batch_param)
        param_btn_layout.addWidget(self.btn_remove_batch_param)
        params_layout.addWidget(self.batch_params_table)
        params_layout.addLayout(param_btn_layout)
        params_group.setLayout(params_layout)
        layout.addWidget(params_group)
        files_group = QGroupBox("📁 数据文件设置")
        files_layout = QVBoxLayout()
        self.files_info_label = QLabel()
        self._update_files_info_label()
        self.files_info_label.setStyleSheet("QLabel { background-color: #f0f8ff; padding: 10px; border-radius: 5px; }")
        files_layout.addWidget(self.files_info_label)
        pattern_layout = QHBoxLayout()
        pattern_layout.addWidget(QLabel("文件命名模式："))
        self.file_pattern_edit = QLineEdit("data/exp_{index}")
        self.file_pattern_edit.setToolTip("使用 {index} 作为序号占位符，如：data/exp_{index}.csv")
        pattern_layout.addWidget(self.file_pattern_edit)
        files_layout.addLayout(pattern_layout)
        self.btn_batch_select_files = QPushButton("📂 批量选择数据文件")
        self.btn_batch_select_files.clicked.connect(self._batch_select_files)
        files_layout.addWidget(self.btn_batch_select_files)
        self.batch_file_list_widget = QListWidget()
        self.batch_file_list_widget.setMaximumHeight(100)
        files_layout.addWidget(self.batch_file_list_widget)
        files_group.setLayout(files_layout)
        layout.addWidget(files_group)
        actions_group = QGroupBox("🚀 操作")
        actions_layout = QVBoxLayout()
        self.btn_run_batch = QPushButton("🧮 运行批量分析")
        self.btn_run_batch.clicked.connect(self._run_batch_analysis)
        self.btn_run_batch.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.btn_export_batch = QPushButton("📊 导出对比表格")
        self.btn_export_batch.clicked.connect(self._export_batch_results)
        self.btn_export_batch.setEnabled(False)
        self.btn_save_batch_config = QPushButton("💾 保存配置")
        self.btn_save_batch_config.clicked.connect(self._save_batch_config)
        self.btn_load_batch_config = QPushButton("📥 加载配置")
        self.btn_load_batch_config.clicked.connect(self._load_batch_config)
        actions_layout.addWidget(self.btn_run_batch)
        actions_layout.addWidget(self.btn_export_batch)
        actions_layout.addWidget(self.btn_save_batch_config)
        actions_layout.addWidget(self.btn_load_batch_config)
        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)
        layout.addStretch()
        self._init_default_batch_params()
        return panel
    def _create_batch_results_panel(self):
        """创建批量实验结果面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        self.batch_results_tabs = QTabWidget()
        table_widget = QWidget()
        table_layout = QVBoxLayout(table_widget)
        
        self.batch_results_table = QTableWidget()
        table_layout.addWidget(QLabel("批量实验结果对比："))
        table_layout.addWidget(self.batch_results_table)
        
        self.batch_results_tabs.addTab(table_widget, "📊 结果对比表")
        
  
        curve_widget = QWidget()
        curve_layout = QVBoxLayout(curve_widget)
        
        self.canvas_batch_efficiency = MatplotlibCanvas(self)
        curve_layout.addWidget(self.canvas_batch_efficiency)
        
        self.batch_results_tabs.addTab(curve_widget, "📈 效率曲线")
        
      
        power_widget = QWidget()
        power_layout = QVBoxLayout(power_widget)
        
        self.canvas_batch_power = MatplotlibCanvas(self)
        power_layout.addWidget(self.canvas_batch_power)
        
        self.batch_results_tabs.addTab(power_widget, "⚡ 功率分析")
        
        layout.addWidget(self.batch_results_tabs)
        
        return panel
    
    def _init_default_batch_params(self):
   
        self._on_batch_explore_type_changed(self.batch_explore_type.currentText())
    
    def _update_files_info_label(self):

        if hasattr(self, 'batch_explore_type'):
            info_text = (
                "文件选择说明（统一单文件模式）：\n"
                "每组实验只需要一个数据文件（使用通道1、2）\n"
                "方式1：使用文件命名模式\n"
                "  - 格式：路径/前缀_{index}.csv\n"
                "  - 示例：data/exp_{index}.csv\n"
                "方式2：批量选择文件（推荐）\n"
                "  - 直接选择所有实验数据文件\n"
                "注：系统只分析发电机输出数据（AIN1、AIN2通道）"
            )
            if hasattr(self, 'files_info_label'):
                self.files_info_label.setText(info_text)
    
    def _on_batch_explore_type_changed(self, explore_type):
   
        self.batch_params_table.setRowCount(0)
        if explore_type == "输入电压影响":
            self.fixed_params_label.setText("分析通道：AIN1、AIN2（发电机输出）")
            self.batch_params_table.setHorizontalHeaderLabels(["序号", "输入电压(V)", "输入功率(W)"])
   
            default_voltages = [
                (10.0, 9.0),
                (10.5, 11.0),
                (11.0, 12.0),
                (11.5, 12.0),
                (12.0, 13.0),
                (12.5, 14.0)
            ]
            self.batch_params_table.setRowCount(len(default_voltages))
            for i, (voltage, power) in enumerate(default_voltages):
                self.batch_params_table.setItem(i, 0, QTableWidgetItem(str(i+1)))
                self.batch_params_table.setItem(i, 1, QTableWidgetItem(str(voltage)))
                self.batch_params_table.setItem(i, 2, QTableWidgetItem(str(power)))
        elif explore_type == "负载电阻影响":
            self.fixed_params_label.setText("分析通道：AIN1、AIN2（发电机输出）")
            self.batch_params_table.setHorizontalHeaderLabels(["序号", "负载电阻(Ω)", "输入功率(W)"])
 
            default_resistances = [
                (2.5, 11.0),
                (3.0, 12.0),
                (3.5, 13.0),
                (4.0, 13.0),
                (4.5, 14.0)
            ]
            self.batch_params_table.setRowCount(len(default_resistances))
            for i, (resistance, power) in enumerate(default_resistances):
                self.batch_params_table.setItem(i, 0, QTableWidgetItem(str(i+1)))
                self.batch_params_table.setItem(i, 1, QTableWidgetItem(str(resistance)))
                self.batch_params_table.setItem(i, 2, QTableWidgetItem(str(power)))
        else: 
            self.fixed_params_label.setText("分析通道：AIN1、AIN2（发电机输出）")
            self.batch_params_table.setHorizontalHeaderLabels(["序号", "磁场距离(mm)", "输入功率(W)"])
            default_distances = [
                (0.0, 10.0),
                (10.0, 10.0),
                (20.0, 10.0),
                (30.0, 10.0),
                (40.0, 10.0),
                (50.0, 10.0)
            ] 
            self.batch_params_table.setRowCount(len(default_distances))
            for i, (distance, power) in enumerate(default_distances):
                self.batch_params_table.setItem(i, 0, QTableWidgetItem(str(i+1)))
                self.batch_params_table.setItem(i, 1, QTableWidgetItem(str(distance)))
                self.batch_params_table.setItem(i, 2, QTableWidgetItem(str(power)))
        
       
        self._update_files_info_label()
    
    def _add_batch_param(self):

        row_count = self.batch_params_table.rowCount()
        self.batch_params_table.insertRow(row_count)
        self.batch_params_table.setItem(row_count, 0, QTableWidgetItem(str(row_count+1)))
        self.batch_params_table.setItem(row_count, 1, QTableWidgetItem("0.0"))
        self.batch_params_table.setItem(row_count, 2, QTableWidgetItem("0.0"))
    
    def _remove_batch_param(self):
    
        current_row = self.batch_params_table.currentRow()
        if current_row >= 0:
            self.batch_params_table.removeRow(current_row)
      
            for i in range(self.batch_params_table.rowCount()):
                self.batch_params_table.setItem(i, 0, QTableWidgetItem(str(i+1)))
    
    def _batch_select_files(self):

        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("CSV 文件 (*.csv)")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        
        if file_dialog.exec():
            files = file_dialog.selectedFiles()
            self.batch_file_list.clear()
            self.batch_file_list_widget.clear()
 
            for i, file in enumerate(sorted(files)):
                self.batch_file_list.append(file) 
                self.batch_file_list_widget.addItem(f"组{i+1}: {os.path.basename(file)}")
            
            self.log(f"已选择 {len(self.batch_file_list)} 个数据文件", "SUCCESS")
    
    def _run_batch_analysis(self):

        self.batch_config = ExperimentConfig(is_factor_exploration_mode=True)
        
     
        params = self._get_batch_params_from_table()
        if not params:
            QMessageBox.warning(self, "参数错误", "请正确填写实验参数")
            return
    
        if self.batch_explore_type.currentText() == "输入电压影响":
            voltage_levels = []
            for voltage, power in params:
                voltage_levels.append({'drive_v': voltage, 'power_input': power})
            self.batch_config.configure_voltage_exploration(voltage_levels, r_load_fixed=10.0)
        elif self.batch_explore_type.currentText() == "负载电阻影响":
            resistance_power_levels = []
            for r_load, power in params:
                resistance_power_levels.append({'r_load': r_load, 'power_input': power})
            self.batch_config.configure_resistance_exploration(resistance_power_levels, drive_v_fixed=12.0)
        else:  # 磁场距离影响
       
            distance_power_pairs = []
            for distance, power in params:
                distance_power_pairs.append((distance, power))
            
          
            self.batch_config.exploration_type = 'magnetic_distance'
            self.batch_config.fixed_params = {
                'drive_v': 12.0,
                'r_load': 10.0
            }
            self.batch_config.variable_params = []
            for distance, power in distance_power_pairs:
                self.batch_config.variable_params.append({
                    'magnetic_distance': distance,
                    'power_input': power
                })
        
      
        for key in ["reference_v", "initial_v", "sampling_freq", "r_load"]:
            if key in self.param_inputs:
                self.batch_config.common_params[key] = float(self.param_inputs[key].text())
        
       
        self.batch_analyzer = BatchExperimentAnalyzer(self.batch_config)
        
        self.log("开始批量实验分析(因素探究模式)...", "INFO")
        
        try:
           
            base_calc_params = {}
            required_dual_params = ["reference_v", "initial_v", "r_load", "sampling_freq"]
          

            for key in required_dual_params:
                if key in self.param_inputs and self.param_inputs[key].text():
                    try:
                        base_calc_params[key] = float(self.param_inputs[key].text().replace(",", "."))
                    except ValueError:
                        QMessageBox.warning(self, "参数错误", f"双机标定参数 '{key}' 无效，请检查！")
                        self.log(f"错误: 双机标定参数 '{key}' 无效", "ERROR")
                        return
                else:
                    QMessageBox.warning(self, "参数缺失", f"双机标定参数 '{key}' 未填写，请检查！")
                    self.log(f"错误: 双机标定参数 '{key}' 未填写", "ERROR")
                    return

           
            self.batch_config.common_params['r_load'] = base_calc_params['r_load']
            self.batch_config.common_params['reference_v'] = base_calc_params['reference_v']
            self.batch_config.common_params['initial_v'] = base_calc_params['initial_v']
            self.batch_config.common_params['sampling_freq'] = base_calc_params['sampling_freq']


           
            files_to_process = []
            if self.batch_file_list:
                for file_path_or_pair in self.batch_file_list:
                    current_file = file_path_or_pair if isinstance(file_path_or_pair, str) else file_path_or_pair[0]
                    files_to_process.append(current_file)
            else:
                pattern = self.file_pattern_edit.text()
                for i in range(len(params)):
                    files_to_process.append(pattern.format(index=i+1) + ".csv")
            
            if len(files_to_process) != len(params):
                QMessageBox.warning(self, "文件/参数不匹配", f"需要处理的文件数 ({len(files_to_process)}) 与参数组数 ({len(params)}) 不一致。")
                self.log(f"错误: 文件数与参数组数不匹配", "ERROR")
                return

            for i, file_path in enumerate(files_to_process):
                if not os.path.exists(file_path):
                    print(f"警告: 文件 {file_path} 未找到，跳过组 {i+1}")
                    self.log(f"警告: 文件 {file_path} 未找到，跳过组 {i+1}", "WARNING")
                    continue

            
                variable_value, power_input_from_table = params[i]
                
                current_run_params = base_calc_params.copy()
                current_run_params['power_input'] = power_input_from_table 
                
               
                if self.batch_config.exploration_type == 'voltage':
                    current_run_params['drive_v'] = variable_value
                elif self.batch_config.exploration_type == 'resistance':
                    
                    current_run_params['r_load'] = variable_value 
                    current_run_params['drive_v'] = self.batch_config.fixed_params.get('drive_v', 12.0) 
                elif self.batch_config.exploration_type == 'magnetic_distance':
            
                    current_run_params['drive_v'] = self.batch_config.fixed_params.get('drive_v', 12.0)
           
                temp_variable_param = {}
                if self.batch_config.exploration_type == 'voltage':
                    temp_variable_param['drive_v'] = variable_value
                elif self.batch_config.exploration_type == 'resistance':
                    temp_variable_param['r_load'] = variable_value
                elif self.batch_config.exploration_type == 'magnetic_distance':
                    temp_variable_param['magnetic_distance'] = variable_value
                temp_variable_param['power_input'] = power_input_from_table
                
            
                if i >= len(self.batch_config.variable_params):
                    self.batch_config.variable_params.append(temp_variable_param)
                else:
                    self.batch_config.variable_params[i] = temp_variable_param
                
               
                p_config = self.batch_config.get_experiment_params(i)
              
                if self.batch_config.exploration_type == 'resistance':
                     p_config['r_load'] = variable_value
                else:
                     p_config['r_load'] = base_calc_params['r_load'] 

                print(f"\n[RUN] 第 {i+1} 组: {self.batch_config.exploration_type}")
                print(f"[RUN] 文件: {file_path}")
                print(f"[RUN] 最终计算参数: {p_config}")

                calc_result = calculate_unified_efficiencies(
                    zheng_file_path=file_path,
                    fan_file_path=file_path, 
                    reference_v=p_config['reference_v'],
                    initial_v=p_config['initial_v'],
                    r_load=p_config['r_load'], 
                    drive_v=p_config.get('drive_v', 0),
                    power_input=p_config['power_input'],
                    sampling_freq=p_config['sampling_freq']
                )
                if calc_result:
                    factor_efficiency = calc_result["verification"]["finished_efficiency"]
                    plot_data_source = calc_result["verification"]["zheng"]["plot_data"] 
                    avg_power_val = 0
                    max_power_val = 0
                    if plot_data_source and len(plot_data_source["power"]) > 0:
                        avg_power_val = np.mean(plot_data_source["power"])
                        max_power_val = np.max(plot_data_source["power"])
                    
                    simplified_result = {
                        'experiment_params': p_config, 
                        'experiment_index': i + 1,
                        'factor_exploration_mode': True, 
                        'efficiency': factor_efficiency,
                        'plot_data': plot_data_source, 
                        'avg_output_power': avg_power_val,
                        'max_output_power': max_power_val,
                    }
                    self.batch_analyzer.results.append(simplified_result)
            
            if self.batch_analyzer.results:
                self._update_batch_results()
                self.btn_export_batch.setEnabled(True)
                self.log("批量分析完成！", "SUCCESS")
            else:
                QMessageBox.warning(self, "分析失败", "未获得有效结果，请检查数据文件")
                self.log("批量分析失败：无有效结果", "ERROR")
                
        except Exception as e:
            QMessageBox.critical(self, "执行错误", f"批量分析过程中发生错误: {e}")
            self.log(f"批量分析异常: {e}", "ERROR")
    
    def _get_batch_params_from_table(self):
       
        params = []
        for row in range(self.batch_params_table.rowCount()):
            try:
                value = float(self.batch_params_table.item(row, 1).text())
                power = float(self.batch_params_table.item(row, 2).text())
                params.append((value, power))
            except:
                return None
        return params
    
    def _update_batch_results(self):
     
        if not self.batch_analyzer or not self.batch_analyzer.results:
            return
        
        results = self.batch_analyzer.results
        self.batch_results_table.setRowCount(len(results))
        
      
        is_factor_mode = results[0].get('factor_exploration_mode', False) if results else False

        if is_factor_mode:
           
            self.batch_results_table.setColumnCount(7)
            col_labels = ["实验组", "变量值", "输入功率(W)", "效率(%)", "平均输出功率(W)", "最大输出功率(W)", "相对基准(%)"]
            if self.batch_config.exploration_type == 'voltage':
                col_labels[1] = "输入电压(V)"
            elif self.batch_config.exploration_type == 'resistance':
                col_labels[1] = "负载电阻(Ω)"
            else: # magnetic_distance
                col_labels[1] = "磁场距离(mm)"
            self.batch_results_table.setHorizontalHeaderLabels(col_labels)
        else:
           
            self.batch_results_table.setColumnCount(11) 
            self.batch_results_table.setHorizontalHeaderLabels([
                "实验组", "输入电压(V)", "输入功率(W)",
                "验证-正接效率(%)", "验证-反接效率(%)", "验证-综合效率(%)",
                "理论-正接效率(%)", "理论-反接效率(%)", "理论-综合效率(%)",
                "综合效率差异(%)", "相对误差(%)"
            ])

        base_efficiency = 0
        if results:
            if is_factor_mode:
                base_efficiency = results[0].get('efficiency', 0)
            else:
                base_efficiency = results[0].get('verification', {}).get('finished_efficiency', 0)
        
        for i, result in enumerate(results):
            params = result['experiment_params']
            col = 0
            self.batch_results_table.setItem(i, col, QTableWidgetItem(str(result['experiment_index'])))
            col += 1

            if is_factor_mode:
                
                if self.batch_config.exploration_type == 'voltage':
                    self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{params.get('drive_v', 0):.1f}"))
                elif self.batch_config.exploration_type == 'resistance':
                    self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{params.get('r_load', 0):.1f}"))
                elif 'magnetic_distance' in params: 
                    self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{params['magnetic_distance']:.1f}"))
                else:
                     self.batch_results_table.setItem(i, col, QTableWidgetItem("N/A"))
                col += 1
             
                self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{params['power_input']:.1f}"))
                col += 1
           
                efficiency = result.get('efficiency', 0)
                self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{efficiency*100:.2f}"))
                col += 1
                self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{result.get('avg_output_power', 0):.2f}"))
                col += 1
                self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{result.get('max_output_power', 0):.2f}"))
                col += 1
                if base_efficiency > 0:
                    relative = (efficiency / base_efficiency - 1) * 100
                    self.batch_results_table.setItem(i, col, QTableWidgetItem(f"{relative:+.1f}"))
                else:
                    self.batch_results_table.setItem(i, col, QTableWidgetItem("--"))
            else:
              
                for j in range(col, self.batch_results_table.columnCount()):
                     self.batch_results_table.setItem(i, j, QTableWidgetItem("--"))

        self._update_batch_plots()
    
    def _update_batch_plots(self):
       
        if not self.batch_analyzer or not self.batch_analyzer.results:
            return
        
        results = self.batch_analyzer.results
        is_factor_mode = results[0].get('factor_exploration_mode', False) if results else False
        
      
        x_values = []
        efficiencies = []
        avg_powers = []
        x_axis_labels_for_bar_chart = []
        
        x_label_plot = ""
        title_prefix_plot = ""

        for i, result in enumerate(results):
            params = result['experiment_params']
            current_x_val = 0
            current_x_tick_label = ""

            if self.batch_config.exploration_type == 'voltage':
                current_x_val = params.get('drive_v', 0)
                current_x_tick_label = f'{current_x_val:.1f}V'
                x_label_plot = '输入电压 (V)'
                title_prefix_plot = '电压'
            elif self.batch_config.exploration_type == 'resistance':
                current_x_val = params.get('r_load', 0)
                current_x_tick_label = f'{current_x_val:.1f}Ω'
                x_label_plot = '负载电阻 (Ω)'
                title_prefix_plot = '负载电阻'
            elif self.batch_config.exploration_type == 'magnetic_distance':
                current_x_val = params.get('magnetic_distance', 0)
                current_x_tick_label = f'{current_x_val:.0f}mm'
                x_label_plot = '磁场距离 (mm)'
                title_prefix_plot = '磁场距离'
            
            x_values.append(current_x_val)
            x_axis_labels_for_bar_chart.append(current_x_tick_label)
            efficiencies.append(result.get('efficiency', 0) * 100)
            avg_powers.append(result.get('avg_output_power', 0))
        
   
        self.canvas_batch_efficiency.axes.cla()
        if x_values and efficiencies:
            self.canvas_batch_efficiency.axes.plot(x_values, efficiencies, 'o-', label='效率', 
                                                   markersize=8, linewidth=1.5, color='dodgerblue')
            if efficiencies:
                max_idx = np.argmax(efficiencies)
                self.canvas_batch_efficiency.axes.scatter([x_values[max_idx]], [efficiencies[max_idx]], 
                                                         s=150, c='red', marker='*', 
                                                         label=f'最高效率: {efficiencies[max_idx]:.2f}%', zorder=5)
            for x, y in zip(x_values, efficiencies):
                self.canvas_batch_efficiency.axes.annotate(f'{y:.1f}', (x, y), 
                                                          textcoords="offset points", xytext=(0,10), ha='center', fontsize=8)
        
        self.canvas_batch_efficiency.axes.set_xlabel(x_label_plot)
        self.canvas_batch_efficiency.axes.set_ylabel('效率 (%)')
        self.canvas_batch_efficiency.axes.set_title(f'效率随{title_prefix_plot}变化曲线')
        self.canvas_batch_efficiency.axes.legend()
        self.canvas_batch_efficiency.axes.grid(True, alpha=0.4, linestyle='--')
        self.canvas_batch_efficiency.draw()
        
       
        self.canvas_batch_power.axes.cla()
        if x_values and avg_powers:
            bar_indices = np.arange(len(x_values))
            bars = self.canvas_batch_power.axes.bar(bar_indices, avg_powers, 
                                                    color='skyblue', alpha=0.8, width=0.6)
            for i, bar in enumerate(bars):
                yval = bar.get_height()
                self.canvas_batch_power.axes.text(bar.get_x() + bar.get_width()/2.0, yval + max(avg_powers)*0.01, 
                                                 f'{yval:.2f}W', ha='center', va='bottom', fontsize=8)
            self.canvas_batch_power.axes.set_xticks(bar_indices)
            self.canvas_batch_power.axes.set_xticklabels(x_axis_labels_for_bar_chart, rotation=30, ha='right')
        
        self.canvas_batch_power.axes.set_xlabel(title_prefix_plot)
        self.canvas_batch_power.axes.set_ylabel('平均输出功率 (W)')
        self.canvas_batch_power.axes.set_title(f'不同{title_prefix_plot}下的平均输出功率')
        self.canvas_batch_power.axes.grid(True, alpha=0.4, axis='y', linestyle='--')
        plt.setp(self.canvas_batch_power.axes.get_xticklabels(), fontsize=9)
        self.canvas_batch_power.fig.tight_layout()
        self.canvas_batch_power.draw()
    
    def _export_batch_results(self):
       
        if self.batch_analyzer:
            df = self.batch_analyzer.generate_comparison_table()
            if df is not None:
                QMessageBox.information(self, "导出成功", "批量实验结果已导出到Excel文件")
    
    def _save_batch_config(self):
      
        if self.batch_config:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存配置", "batch_config.json", "JSON文件 (*.json)"
            )
            if file_path:
                self.batch_config.save_config(file_path)
                self.log(f"配置已保存至: {file_path}", "SUCCESS")
    
    def _load_batch_config(self):
       
        file_path, _ = QFileDialog.getOpenFileName(
            self, "加载配置", "", "JSON文件 (*.json)"
        )
        if file_path:
            try:
                self.batch_config = ExperimentConfig.load_config(file_path)
                # TODO: 更新界面显示
                self.log(f"配置已加载: {file_path}", "SUCCESS")
            except Exception as e:
                QMessageBox.critical(self, "加载失败", f"加载配置失败: {e}")



    def _create_control_panel(self):
     
        panel = QWidget()
        panel.setFixedWidth(420)
        layout = QVBoxLayout(panel)
        
   
        file_group = QGroupBox("📁 数据文件导入")
        file_layout = QVBoxLayout()
        
   
        info_label = QLabel(
            "通道分配说明：\n"
            "• 验证实验：AIN1-2 (通道1-2)\n"
            "• 理论实验：AIN5-7 (通道5-7)"
        )
        info_label.setStyleSheet("QLabel { background-color: #f0f8ff; padding: 10px; border-radius: 5px; }")
        file_layout.addWidget(info_label)
        
        self.zheng_label = QLabel("未选择正接数据文件")
        self.btn_load_zheng = QPushButton("📊 导入正接数据")
        self.btn_load_zheng.clicked.connect(lambda: self._load_file('zheng'))
        
        self.fan_label = QLabel("未选择反接数据文件")
        self.btn_load_fan = QPushButton("📊 导入反接数据")
        self.btn_load_fan.clicked.connect(lambda: self._load_file('fan'))
        
        file_layout.addWidget(self.btn_load_zheng)
        file_layout.addWidget(self.zheng_label)
        file_layout.addWidget(self.btn_load_fan)
        file_layout.addWidget(self.fan_label)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
     
        params_group = QGroupBox("⚙️ 实验参数设置")
        params_layout = QVBoxLayout()
        
        self.param_inputs = {}
        param_info = {
            "reference_v": ("电流比例值 (A/V):", "电流测量模块的校准参数"),
            "initial_v": ("基准电压 (V):", "电流测量模块的零点偏置"),
            "r_load": ("负载电阻 R (Ω):", "发电机负载电阻"),
            "drive_v": ("驱动电压 (V):", "理论实验用，驱动电机电压"),
            "power_input": ("平均输入功率 (W):", "验证实验用，扭矩传感器测量值"),
            "sampling_freq": ("采样频率 (Hz):", "数据采集频率")
        }
        
        double_validator = QDoubleValidator()
        double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        
        for key, (label_text, tooltip) in param_info.items():
            label = QLabel(label_text)
            line_edit = QLineEdit(self.default_params.get(key, ""))
            line_edit.setValidator(double_validator)
            line_edit.setToolTip(tooltip)
            self.param_inputs[key] = line_edit
            row = QHBoxLayout()
            row.addWidget(label)
            row.addWidget(line_edit)
            params_layout.addLayout(row)
            
  
        data_points_label = QLabel("数据点处理 (可选):")
        data_points_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        params_layout.addWidget(data_points_label)
        
        for key, label_text in [("points_to_process_zheng", "正接处理前N个点:"), 
                                ("points_to_process_fan", "反接处理前N个点:")]:
            label = QLabel(label_text)
            line_edit = QLineEdit()
            line_edit.setPlaceholderText("留空处理全部数据")
            self.param_inputs[key] = line_edit
            row = QHBoxLayout()
            row.addWidget(label)
            row.addWidget(line_edit)
            params_layout.addLayout(row)
        
        params_group.setLayout(params_layout)
        layout.addWidget(params_group)
        

        actions_group = QGroupBox("🚀 操作")
        actions_layout = QVBoxLayout()
        
        self.btn_calculate = QPushButton("🧮 开始统一计算")
        self.btn_calculate.clicked.connect(self._calculate)
        self.btn_calculate.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        
        self.btn_export = QPushButton("📄 导出分析报告")
        self.btn_export.clicked.connect(self._export_report)
        self.btn_export.setEnabled(False)
        
        self.btn_principle = QPushButton("📖 查看原理说明")
        self.btn_principle.clicked.connect(self._show_principle)
        
        actions_layout.addWidget(self.btn_calculate)
        actions_layout.addWidget(self.btn_export)
        actions_layout.addWidget(self.btn_principle)
        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)
        
        layout.addStretch()
        return panel

    def _create_results_panel(self):
 
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
    
        self.results_tabs = QTabWidget()
        
      
        efficiency_tab = self._create_efficiency_tab()
        self.results_tabs.addTab(efficiency_tab, "📊 效率结果")
        
      
        stats_tab = self._create_stats_tab()
        self.results_tabs.addTab(stats_tab, "📈 数据统计")
        
     
        plots_tab = self._create_plots_tab()
        self.results_tabs.addTab(plots_tab, "📉 图表分析")
        
        layout.addWidget(self.results_tabs)
        return panel

    def _create_efficiency_tab(self):
    
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
   
        self.efficiency_table = QTableWidget()
        self.efficiency_table.setRowCount(4)
        self.efficiency_table.setColumnCount(5)
        self.efficiency_table.setHorizontalHeaderLabels(["项目", "验证实验", "理论实验", "差异", "差异率(%)"])
        self.efficiency_table.setVerticalHeaderLabels(["正向效率", "反向效率", "综合效率", "计算方法"])
        
   
        self.efficiency_table.horizontalHeader().setStretchLastSection(True)
        self.efficiency_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        
        self._init_efficiency_table()
        
        layout.addWidget(QLabel("效率计算结果对比："))
        layout.addWidget(self.efficiency_table)
        
   
        self.analysis_text = QTextEdit()
        self.analysis_text.setReadOnly(True)
        self.analysis_text.setMaximumHeight(150)
        self.analysis_text.setPlainText("等待计算结果...")
        
        layout.addWidget(QLabel("差异分析："))
        layout.addWidget(self.analysis_text)
        
        return widget

    def _create_stats_tab(self):
   
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
    
        self.stats_table = QTableWidget()
        self.stats_table.setRowCount(8)  
        self.stats_table.setColumnCount(7) 
        self.stats_table.setHorizontalHeaderLabels([
            "通道", 
            "正接-最大值", "正接-最小值", "正接-平均值",
            "反接-最大值", "反接-最小值", "反接-平均值"
        ])
        
       
        for i in range(8):
            self.stats_table.setItem(i, 0, QTableWidgetItem(f"AIN{i+1}"))
        
            if i+1 in [1, 2]: 
                self.stats_table.item(i, 0).setBackground(Qt.GlobalColor.lightGray)
            elif i+1 in [5, 6, 7]: 
                self.stats_table.item(i, 0).setBackground(Qt.GlobalColor.lightGray)
        
        layout.addWidget(QLabel("多通道数据统计："))
        layout.addWidget(self.stats_table)
        
        return widget

    def _create_plots_tab(self):
      
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        
        plot_tabs = QTabWidget()
        
       
        self.canvas_current = MatplotlibCanvas(self)
        plot_tabs.addTab(self.canvas_current, "电流对比")
        
      
        self.canvas_power = MatplotlibCanvas(self)
        plot_tabs.addTab(self.canvas_power, "功率对比")
        
   
        self.canvas_efficiency_bar = MatplotlibCanvas(self)
        plot_tabs.addTab(self.canvas_efficiency_bar, "效率对比")
        
        layout.addWidget(plot_tabs)
        return widget

    def _init_efficiency_table(self):
        for i in range(4):
            for j in range(1, 5):
                self.efficiency_table.setItem(i, j, QTableWidgetItem("--"))
        
        self.efficiency_table.setItem(3, 0, QTableWidgetItem("计算方法"))
        self.efficiency_table.setItem(3, 1, QTableWidgetItem("∫I²R / (P×t)"))
        self.efficiency_table.setItem(3, 2, QTableWidgetItem("∫I²R / ∫VI"))
        self.efficiency_table.setItem(3, 3, QTableWidgetItem("--"))
        self.efficiency_table.setItem(3, 4, QTableWidgetItem("--"))

    def _load_file(self, file_type):

        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("CSV 文件 (*.csv)")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        if file_dialog.exec():
            file_path = file_dialog.selectedFiles()[0]
            if file_type == 'zheng':
                self.zheng_file = file_path
                self.zheng_label.setText(f"已选择: {file_path.split('/')[-1]}")
                self.log(f"已加载正接数据文件: {file_path}", "SUCCESS")
            else:
                self.fan_file = file_path
                self.fan_label.setText(f"已选择: {file_path.split('/')[-1]}")
                self.log(f"已加载反接数据文件: {file_path}", "SUCCESS")

    def _validate_params(self):

        params = {}
        try:
            for key in self.param_inputs:
                if key.startswith("points_to_process"):
                    text_value = self.param_inputs[key].text().strip()
                    if not text_value:
                        params[key] = None
                    else:
                        value = int(text_value)
                        if value < 0:
                            raise ValueError(f"处理点数不能为负")
                        params[key] = value
                else:
                    text_value = self.param_inputs[key].text().replace(",", ".")
                    if not text_value:
                        raise ValueError(f"参数 '{key}' 不能为空")
                    value = float(text_value)
                    if key in ["r_load", "sampling_freq"] and value <= 0:
                        raise ValueError(f"参数 '{key}' 必须为正数")
                    params[key] = value
            return params
        except ValueError as e:
            QMessageBox.warning(self, "参数错误", f"参数输入无效: {e}")
            self.log(f"参数验证失败: {e}", "ERROR")
            return None

    def _calculate(self):

        if not self.zheng_file or not self.fan_file:
            QMessageBox.warning(self, "文件未选择", "请先导入正接和反接数据文件")
            self.log("计算失败：未选择数据文件", "WARNING")
            return

        params = self._validate_params()
        if not params:
            return

        self.log("开始统一计算...", "INFO")
        
        try:
            self.results = calculate_unified_efficiencies(
                zheng_file_path=self.zheng_file,
                fan_file_path=self.fan_file,
                reference_v=params["reference_v"],
                initial_v=params["initial_v"],
                r_load=params["r_load"],
                drive_v=params["drive_v"],
                power_input=params["power_input"],
                sampling_freq=params["sampling_freq"],
                points_to_process_zheng=params["points_to_process_zheng"],
                points_to_process_fan=params["points_to_process_fan"]
            )

            if self.results:
                self._update_results()
                self.btn_export.setEnabled(True)
                self.log("计算完成！", "SUCCESS")
                QMessageBox.information(self, "计算完成", "统一效率计算已完成，请查看结果")
            else:
                QMessageBox.critical(self, "计算错误", "计算失败，请检查数据文件和参数")
                self.log("计算失败：返回结果为空", "ERROR")

        except Exception as e:
            QMessageBox.critical(self, "执行错误", f"计算过程中发生错误: {e}")
            self.log(f"计算异常: {e}", "ERROR")

    def _update_results(self):
 
        if not self.results:
            return
        
        self._update_efficiency_table()
        
        self._update_stats_table()
        

        self._update_plots()
        

        self._update_analysis()

    def _update_efficiency_table(self):
        ver = self.results["verification"]
        theo = self.results["theoretical"]
        comp = self.results["comparison"]
        
   
        self.efficiency_table.setItem(0, 0, QTableWidgetItem("正向效率"))
        self.efficiency_table.setItem(0, 1, QTableWidgetItem(f"{ver['zheng']['efficiency']*100:.3f}%"))
        self.efficiency_table.setItem(0, 2, QTableWidgetItem(f"{theo['zheng']['efficiency']*100:.3f}%"))
        self.efficiency_table.setItem(0, 3, QTableWidgetItem(f"{comp['zheng_diff']*100:.3f}%"))
        diff_rate = comp['zheng_diff'] / ver['zheng']['efficiency'] * 100 if ver['zheng']['efficiency'] > 0 else 0
        self.efficiency_table.setItem(0, 4, QTableWidgetItem(f"{diff_rate:.1f}%"))
        

        self.efficiency_table.setItem(1, 0, QTableWidgetItem("反向效率"))
        self.efficiency_table.setItem(1, 1, QTableWidgetItem(f"{ver['fan']['efficiency']*100:.3f}%"))
        self.efficiency_table.setItem(1, 2, QTableWidgetItem(f"{theo['fan']['efficiency']*100:.3f}%"))
        self.efficiency_table.setItem(1, 3, QTableWidgetItem(f"{comp['fan_diff']*100:.3f}%"))
        diff_rate = comp['fan_diff'] / ver['fan']['efficiency'] * 100 if ver['fan']['efficiency'] > 0 else 0
        self.efficiency_table.setItem(1, 4, QTableWidgetItem(f"{diff_rate:.1f}%"))
        

        self.efficiency_table.setItem(2, 0, QTableWidgetItem("综合效率"))
        self.efficiency_table.setItem(2, 1, QTableWidgetItem(f"{ver['finished_efficiency']*100:.3f}%"))
        self.efficiency_table.setItem(2, 2, QTableWidgetItem(f"{theo['finished_efficiency']*100:.3f}%"))
        self.efficiency_table.setItem(2, 3, QTableWidgetItem(f"{comp['finished_diff']*100:.3f}%"))
        diff_rate = comp['finished_diff'] / ver['finished_efficiency'] * 100 if ver['finished_efficiency'] > 0 else 0
        self.efficiency_table.setItem(2, 4, QTableWidgetItem(f"{diff_rate:.1f}%"))

    def _update_stats_table(self):
        ver_zheng_stats = self.results["verification"]["zheng"]["stats"]
        ver_fan_stats = self.results["verification"]["fan"]["stats"]
        
        for i in range(8):
            channel = f"AIN{i+1}"
            
            if channel in ver_zheng_stats:
                stats = ver_zheng_stats[channel]
                self.stats_table.setItem(i, 1, QTableWidgetItem(f"{stats['max']:.4f}" if not np.isnan(stats['max']) else "N/A"))
                self.stats_table.setItem(i, 2, QTableWidgetItem(f"{stats['min']:.4f}" if not np.isnan(stats['min']) else "N/A"))
                self.stats_table.setItem(i, 3, QTableWidgetItem(f"{stats['avg']:.4f}" if not np.isnan(stats['avg']) else "N/A"))
            
            if channel in ver_fan_stats:
                stats = ver_fan_stats[channel]
                self.stats_table.setItem(i, 4, QTableWidgetItem(f"{stats['max']:.4f}" if not np.isnan(stats['max']) else "N/A"))
                self.stats_table.setItem(i, 5, QTableWidgetItem(f"{stats['min']:.4f}" if not np.isnan(stats['min']) else "N/A"))
                self.stats_table.setItem(i, 6, QTableWidgetItem(f"{stats['avg']:.4f}" if not np.isnan(stats['avg']) else "N/A"))

    def _update_plots(self):
        ver = self.results["verification"]
        theo = self.results["theoretical"]
        
        datasets = []
        if len(ver["zheng"]["plot_data"]["time"]) > 0:
            datasets.append((ver["zheng"]["plot_data"]["time"], ver["zheng"]["plot_data"]["current"], "验证-正接"))
        if len(ver["fan"]["plot_data"]["time"]) > 0:
            datasets.append((ver["fan"]["plot_data"]["time"], ver["fan"]["plot_data"]["current"], "验证-反接"))
        if len(theo["zheng"]["plot_data"]["time"]) > 0:
            datasets.append((theo["zheng"]["plot_data"]["time"], theo["zheng"]["plot_data"]["output_current"], "理论-正接输出"))
        if len(theo["fan"]["plot_data"]["time"]) > 0:
            datasets.append((theo["fan"]["plot_data"]["time"], theo["fan"]["plot_data"]["output_current"], "理论-反接输出"))
        
        self.canvas_current.plot_comparison(datasets, "电流对比", "时间 (s)", "电流 (A)")
        
        datasets = []
        if len(ver["zheng"]["plot_data"]["time"]) > 0:
            datasets.append((ver["zheng"]["plot_data"]["time"], ver["zheng"]["plot_data"]["power"], "验证-正接"))
        if len(ver["fan"]["plot_data"]["time"]) > 0:
            datasets.append((ver["fan"]["plot_data"]["time"], ver["fan"]["plot_data"]["power"], "验证-反接"))
        if len(theo["zheng"]["plot_data"]["time"]) > 0:
            datasets.append((theo["zheng"]["plot_data"]["time"], theo["zheng"]["plot_data"]["output_power"], "理论-正接"))
        if len(theo["fan"]["plot_data"]["time"]) > 0:
            datasets.append((theo["fan"]["plot_data"]["time"], theo["fan"]["plot_data"]["output_power"], "理论-反接"))
        
        self.canvas_power.plot_comparison(datasets, "输出功率对比", "时间 (s)", "功率 (W)")
        
  
        self._plot_efficiency_comparison()

    def _plot_efficiency_comparison(self):
        categories = ['正向效率', '反向效率', '综合效率']
        verification_values = [
            self.results["verification"]["zheng"]["efficiency"] * 100,
            self.results["verification"]["fan"]["efficiency"] * 100,
            self.results["verification"]["finished_efficiency"] * 100
        ]
        theoretical_values = [
            self.results["theoretical"]["zheng"]["efficiency"] * 100,
            self.results["theoretical"]["fan"]["efficiency"] * 100,
            self.results["theoretical"]["finished_efficiency"] * 100
        ]

        self.canvas_efficiency_bar.axes.cla()
        
        x_pos = np.arange(len(categories))
        width = 0.35

        bars1 = self.canvas_efficiency_bar.axes.bar(x_pos - width/2, verification_values, width, 
                                                     label='验证实验', color='skyblue', alpha=0.8)
        bars2 = self.canvas_efficiency_bar.axes.bar(x_pos + width/2, theoretical_values, width,
                                                     label='理论实验', color='lightcoral', alpha=0.8)

        self.canvas_efficiency_bar.axes.set_xlabel('效率类型')
        self.canvas_efficiency_bar.axes.set_ylabel('效率 (%)')
        self.canvas_efficiency_bar.axes.set_title('验证与理论实验效率对比')
        self.canvas_efficiency_bar.axes.set_xticks(x_pos)
        self.canvas_efficiency_bar.axes.set_xticklabels(categories)
        self.canvas_efficiency_bar.axes.legend()
        self.canvas_efficiency_bar.axes.grid(True, alpha=0.3, axis='y')

        for bar in bars1:
            height = bar.get_height()
            self.canvas_efficiency_bar.axes.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                                                f'{height:.1f}%', ha='center', va='bottom', fontsize=9)
        for bar in bars2:
            height = bar.get_height()
            self.canvas_efficiency_bar.axes.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                                                f'{height:.1f}%', ha='center', va='bottom', fontsize=9)

        self.canvas_efficiency_bar.draw()

    def _update_analysis(self):
        comp = self.results["comparison"]
        
        analysis = f"""效率差异分析结果：

1. 正向效率差异: {comp['zheng_diff']*100:.3f}%
   - 验证实验采用平均功率法，假设输入功率恒定
   - 理论实验实时测量输入功率，能捕捉功率波动

2. 反向效率差异: {comp['fan_diff']*100:.3f}%
   - 差异反映了两种测量方法的本质区别
   - 理论方法通常更准确但更复杂

3. 综合效率差异: {comp['finished_diff']*100:.3f}%
   - 综合效率采用几何平均值计算
   - 差异程度反映了实验条件的稳定性

建议：若差异较大(>5%)，应检查实验条件是否稳定。"""
        
        self.analysis_text.setPlainText(analysis)

    def _export_report(self):
        if not self.results:
            return
            
        file_path, _ = QFileDialog.getSaveFileName(self, "导出报告", "unified_analysis_report.txt", "文本文件 (*.txt)")
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("电机效率统一分析报告\n")
                    f.write("=" * 60 + "\n\n")
                    
                    f.write("实验参数：\n")
                    for key, input_widget in self.param_inputs.items():
                        if not key.startswith("points_"):
                            f.write(f"  {key}: {input_widget.text()}\n")
                    f.write("\n")
                    
                    f.write("效率计算结果：\n")
                    f.write("-" * 40 + "\n")
                    f.write("验证实验（平均功率法）：\n")
                    ver = self.results["verification"]
                    f.write(f"  正向效率: {ver['zheng']['efficiency']*100:.3f}%\n")
                    f.write(f"  反向效率: {ver['fan']['efficiency']*100:.3f}%\n")
                    f.write(f"  综合效率: {ver['finished_efficiency']*100:.3f}%\n\n")
                    
                    f.write("理论实验（实时功率法）：\n")
                    theo = self.results["theoretical"]
                    f.write(f"  正向效率: {theo['zheng']['efficiency']*100:.3f}%\n")
                    f.write(f"  反向效率: {theo['fan']['efficiency']*100:.3f}%\n")
                    f.write(f"  综合效率: {theo['finished_efficiency']*100:.3f}%\n\n")
                    
                    f.write("差异分析：\n")
                    comp = self.results["comparison"]
                    f.write(f"  正向效率差异: {comp['zheng_diff']*100:.3f}%\n")
                    f.write(f"  反向效率差异: {comp['fan_diff']*100:.3f}%\n")
                    f.write(f"  综合效率差异: {comp['finished_diff']*100:.3f}%\n")
                    
                self.log(f"报告已导出至: {file_path}", "SUCCESS")
                QMessageBox.information(self, "导出成功", f"报告已保存至:\n{file_path}")
                
            except Exception as e:
                self.log(f"导出失败: {e}", "ERROR")
                QMessageBox.critical(self, "导出失败", f"导出过程中发生错误: {e}")

    def _show_principle(self):
        """显示原理说明"""
        principle_dialog = PrincipleDialog(self)
        principle_dialog.exec()




class PrincipleDialog(QDialog):
    """原理说明对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("原理说明")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(self)
        
        tabs = QTabWidget()
        
        unified_text = """
        <h2>🔬 统一测量系统</h2>
        
        <h3>系统核心设计</h3>
        <p>本系统基于双电机配置，包含一台驱动电机和一台由其带动的发电机（同型号）。可调直流电源为驱动电机供电，驱动电机将电能转化为机械能，再通过联轴器传递给发电机。发电机将接收到的机械能转化为电能输出至负载（如滑动变阻器）。</p>
        <p>关键数据采集点包括：</p>
        <ul>
            <li><strong>驱动电机输入：</strong>通过电流传感器测量驱动电机的输入电流，结合已知的驱动电压，用于计算输入功率。</li>
            <li><strong>发电机输出：</strong>通过电流传感器测量流经负载的输出电流，结合负载电阻，用于计算输出功率。</li>
        </ul>
        <p>数据采集器同步记录这些信号（电流传感器一般将电流信号转换为电压信号后输入采集器）。</p>
        <h3>系统设计理念</h3>
        <p>本系统采用多通道同步采集技术，在一次测量中同时获取驱动电机输入参数和发电机输出参数。这种设计避免了分次测量可能带来的实验条件差异，提高了对比分析的可靠性。</p>

        <h3>通道分配参考</h3>
        <p>具体的通道分配可能因实验设置和所选分析方法而异。以下为本软件在不同实验模式下的一种典型通道用途说明：</p>
        <table border="1" style="border-collapse: collapse; margin: 20px auto; width: 90%;">
            <tr style="background-color: #f0f0f0;">
                <th style="padding: 8px;">通道</th>
                <th style="padding: 8px;">"验证实验"模式下用途</th>
                <th style="padding: 8px;">"理论实验"模式下用途</th>
            </tr>
            <tr><td style="padding: 8px;">发电U：AIN1</td><td style="padding: 8px;">发电机输出信号 (如电流)</td><td style="padding: 8px;">发电机输出信号 (如电流)</td></tr>
            <tr><td style="padding: 8px;">发电I：AIN2</td><td style="padding: 8px;">发电机输出信号 (如电流)</td><td style="padding: 8px;">发电机输出信号 (如电流)</td></tr>
            <tr><td style="padding: 8px;">预留：AIN3-4</td><td style="padding: 8px;" colspan="2">预留或用于其他辅助测量</td></tr>
            <tr><td style="padding: 8px;">发电U：AIN5</td><td style="padding: 8px;">--</td> <td style="padding: 8px;">驱动电机输入信号 (如电流)</td></tr>
            <tr><td style="padding: 8px;">发电I：AIN6</td><td style="padding: 8px;">--</td> <td style="padding: 8px;">驱动电机输入信号 (如电流)</td></tr>
            <tr><td style="padding: 8px;">驱动I：AIN7</td><td style="padding: 8px;">--</td> <td style="padding: 8px;">驱动电机电源电压 (若需测量)</td></tr>
            <tr><td style="padding: 8px;">预留：AIN8</td><td style="padding: 8px;" colspan="2">预留或用于其他辅助测量</td></tr>
        </table>
        <p><em>注意：实际通道用途请根据您的具体实验接线和软件参数配置为准。</em></p>

        <h3>优势分析</h3>
        <ul>
            <li><strong>同步性：</strong>所有关键数据在相同时刻采集，确保实验条件一致性。</li>
            <li><strong>准确性：</strong>减少因分时测量引入的系统误差。</li>
            <li><strong>高效性：</strong>一次实验即可获取进行多种效率分析所需的数据。</li>
        </ul>
        """
        
        verification_text = """
        <h2>⚙️ 单电机效率标定原理 (双机法)</h2>

        <h3>背景与目的</h3>
        <p>在精确评估电机系统性能时，单个电机的自身效率 (η) 是一个关键参数。图片中介绍的"双机标定法"提供了一种创新的实验手段，用于准确测定同型号下单个电机的效率。该方法对于后续分析（如精确计算发电机在特定机械输入下的发电效率）至关重要。</p>

        <h3>双机标定法原理</h3>
        <p>该方法利用两台相同的电机进行组合测试。基本步骤与假设如下：</p>
        <ol>
            <li>将两台相同的电机（设其效率分别为 ηA 和 ηB）机械耦合。由于电机相同，可假定 ηA = ηB = η。</li>
            <li>进行两次组合运行测试：
                <ul>
                    <li><strong>测试1：</strong>电机A作为驱动电机，电机B作为发电机。测得此组合的总效率 η1。此时，η1 = ηA × ηB = η × η = η²。</li>
                    <li><strong>测试2：</strong>电机B作为驱动电机，电机A作为发电机。测得此组合的总效率 η2。此时，η2 = ηB × ηA = η × η = η²。</li>
                </ul>
            </li>
            <li>根据两次测试得到的组合效率 η1 和 η2，单个电机的效率 η 可以通过以下公式计算得出：</li>
        </ol>
        <div style="background-color: #e6ffe6; padding: 15px; margin: 10px 0; text-align: center; font-size: 1.1em;">
            <p><strong>η1 × η2 = η<sup>4</sup></strong></p>
            <p>因此，单个电机效率：</p>
            <p><strong>η = (η1 × η2)<sup>1/4</sup></strong></p>
        </div>

        <h3>应用说明</h3>
        <p>通过双机标定法测得的单电机效率 η，可以用于更精确地计算发电机的机械输入功率，当该电机作为驱动电机使用时。例如，发电机的发电效率 η' 计算公式为：</p>
        <p style="margin-left: 20px;"><strong>η' = P<sub>发电机输出</sub> / P<sub>机械输入</sub></strong></p>
        <p style="margin-left: 20px;">其中，<strong>P<sub>机械输入</sub> = η × P<sub>驱动电机电输入</sub></strong></p>
        <p>此方法提高了复杂系统中能量转换分析的准确度。</p>
        """
        
        theoretical_text = """
        <h2>📐 理论实验原理 (电机-发电机组效率)</h2>

        <h3>基本原理</h3>
        <p>本应用中的"理论实验"旨在计算电机-发电机组的总转换效率。该方法基于严格的物理原理，通过实时同步测量驱动电机的输入电参数和发电机输出至负载的电参数，精确计算瞬时功率并积分得到总输入能量和总输出能量，从而得出整体效率。</p>
        <p>实验设置通常包含一台驱动电机和一台由其机械耦合带动的发电机，如"统一测量系统原理"部分所述。</p>

        <h3>效率计算公式 (电机-发电机组总效率 η<sub>overall</sub>)</h3>
        <p>总效率定义为发电机输出的总有效电能与驱动电机消耗的总电能之比：</p>
        <div style="background-color: #fff0f5; padding: 15px; margin: 10px 0;">
            <p><strong>η<sub>overall</sub> = E<sub>out</sub> / E<sub>in</sub></strong></p>
            <p>其中：</p>
            <p><strong>E<sub>out</sub> (发电机输出总能量) = ∫ P<sub>out</sub>(t) dt = ∫ I<sub>out</sub>²(t) × R<sub>load</sub> dt</strong></p>
            <p><em>(I<sub>out</sub>(t) 是流经负载 R<sub>load</sub> 的瞬时电流)</em></p>
            <p><strong>E<sub>in</sub> (驱动电机输入总能量) = ∫ P<sub>in</sub>(t) dt = ∫ V<sub>drive</sub>(t) × I<sub>in</sub>(t) dt</strong></p>
            <p><em>(V<sub>drive</sub>(t) 是驱动电机两端瞬时电压，I<sub>in</sub>(t) 是流入驱动电机的瞬时电流)</em></p>
        </div>

        <h3>特点</h3>
        <ul>
            <li>基于瞬时值积分，能更准确地反映实际工况下的能量转换。</li>
            <li>需要精确同步测量多个电参数。</li>
            <li>为深入分析电机系统性能提供了可靠依据。</li>
        </ul>
        """
        
        factor_text = """
        <h2>📊 因素探究实验原理</h2>

        <h3>实验目的</h3>
        <p>因素探究实验旨在研究不同物理因素对电磁感应效率的影响规律，通过控制变量法，定量分析各因素与效率之间的关系，为优化电机系统设计提供科学依据。</p>

        <h3>研究因素</h3>
        <ul>
            <li><strong>磁场强度：</strong>通过改变永磁体与电机的距离（0-50mm），研究磁场强度对效率的影响</li>
            <li><strong>温度：</strong>通过控制环境温度或电机温升，研究温度对效率的影响（20-80℃）</li>
            <li><strong>转速：</strong>通过调节驱动电压改变电机转速，研究转速对效率的影响（500-3000rpm）</li>
            <li><strong>负载特性：</strong>通过改变负载电阻值，研究负载对效率的影响（1-50Ω）</li>
        </ul>

        <h3>实验方法</h3>
        <p>采用控制变量法，保持其他条件不变，只改变研究因素：</p>
        <ol>
            <li>设定基准条件，测量基准效率</li>
            <li>逐步改变研究因素的值，在每个值下进行效率测量</li>
            <li>记录不同因素值对应的效率数据</li>
            <li>分析效率随因素变化的规律，找出最优工作点</li>
        </ol>

        <h3>数据分析方法</h3>
        <ul>
            <li><strong>趋势分析：</strong>绘制效率-因素曲线，观察变化趋势</li>
            <li><strong>线性拟合：</strong>计算线性相关系数，判断相关性强弱</li>
            <li><strong>最优值确定：</strong>找出效率最高对应的因素值</li>
            <li><strong>敏感度分析：</strong>计算效率变化率，评估因素影响程度</li>
        </ul>

        <h3>数据采集说明</h3>
        <p>因素探究实验只使用通道AIN1和AIN2的数据：</p>
        <ul>
            <li><strong>AIN1：</strong>发电机输出电压信号</li>
            <li><strong>AIN2：</strong>发电机输出电流信号（经电流采样模块转换）</li>
        </ul>
        <p>通过平均功率法计算效率：η = ∫(I²R)dt / (P_avg × t)</p>

        <h3>实验意义</h3>
        <p>通过因素探究实验，可以：</p>
        <ul>
            <li>深入理解各因素对电磁感应效率的影响机制</li>
            <li>为电机系统优化设计提供定量依据</li>
            <li>指导实际应用中的参数选择和工况优化</li>
            <li>预测不同条件下的系统性能</li>
        </ul>
        """
        
        unified_tab = self._create_principle_tab(unified_text)
        tabs.addTab(unified_tab, "统一测量系统")
        
        verification_tab = self._create_principle_tab(verification_text)
        tabs.addTab(verification_tab, "单电机效率标定 (双机法)")
        
        theoretical_tab = self._create_principle_tab(theoretical_text)
        tabs.addTab(theoretical_tab, "理论实验原理")
        
        factor_tab = self._create_principle_tab(factor_text)
        tabs.addTab(factor_tab, "因素探究实验")
        
        layout.addWidget(tabs)
        
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

    def _create_principle_tab(self, html_text):
        scroll_area = QScrollArea()
        content_label = QLabel(html_text)
        content_label.setWordWrap(True)
        content_label.setTextFormat(Qt.TextFormat.RichText)
        content_label.setAlignment(Qt.AlignmentFlag.AlignTop)
        content_label.setStyleSheet("QLabel { padding: 20px; }")
        
        scroll_area.setWidget(content_label)
        scroll_area.setWidgetResizable(True)
        return scroll_area


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_app = UnifiedMotorAnalysisApp()
    main_app.show()
    sys.exit(app.exec()) 