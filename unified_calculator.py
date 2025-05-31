import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from matplotlib import font_manager
import json
from datetime import datetime
try:
    import openpyxl 
except ImportError:
    print("警告: 未安装openpyxl库，无法导出Excel文件。请运行: pip install openpyxl")


plt.rcParams['font.sans-serif'] = ['SimHei'] 
plt.rcParams['axes.unicode_minus'] = False    

def _calculate_column_stats(column_data_series: pd.Series):
    """Helper function to calculate max, min, avg for a pandas Series."""
    numeric_data = pd.to_numeric(column_data_series, errors='coerce')
    if numeric_data.empty or numeric_data.isnull().all():
        return {"max": np.nan, "min": np.nan, "avg": np.nan}
    return {
        "max": numeric_data.max(),
        "min": numeric_data.min(),
        "avg": numeric_data.mean()
    }

def calculate_simple_efficiency(file_path: str, reference_v: float, initial_v: float, 
                              r_load: float, power_input: float,
                              sampling_freq: float = 87500.0,
                              points_to_process: int | None = None):

    print(f"\n[DEBUG] calculate_simple_efficiency called for file: {file_path}")
    print(f"[DEBUG] Params: ref_v={reference_v}, init_v={initial_v}, r_load={r_load}, power_in={power_input}, samp_freq={sampling_freq}, points={points_to_process}")
    try:
        time_once = 1.0 / sampling_freq
        print(f"[DEBUG] time_once: {time_once:.10f} s")
        
    
        data_df_full = pd.read_csv(file_path, header=None, skiprows=1)
        print(f"[DEBUG] Full data shape: {data_df_full.shape}")
        if points_to_process is not None and points_to_process > 0:
            data_df = data_df_full.head(points_to_process)
        else:
            data_df = data_df_full
        print(f"[DEBUG] Processed data shape: {data_df.shape}")
        
        if data_df.empty or data_df.shape[1] < 3:
            print(f"警告: 数据文件 '{file_path}' 为空或列数不足")
            return None
  
        t_original = pd.to_numeric(data_df.iloc[:, 0], errors='coerce').to_numpy()
        time_array = np.arange(len(t_original)) * time_once
        print(f"[DEBUG] t_original sample (first 5): {t_original[:5]}")
        print(f"[DEBUG] time_array sample (first 5): {time_array[:5]}")
        
        
        output_v_raw = data_df.iloc[:, 2]
        print(f"[DEBUG] output_v_raw (AIN2) sample (first 5): {output_v_raw.head().to_numpy()}")
        output_v = pd.to_numeric(output_v_raw, errors='coerce').to_numpy()
        print(f"[DEBUG] output_v (numeric) sample (first 5): {output_v[:5]}")
        output_i = (output_v - initial_v) / reference_v
        print(f"[DEBUG] output_i (calculated) sample (first 5): {output_i[:5]}")
        
 
        valid_idx = ~np.isnan(output_i) & ~np.isnan(time_array)
        if not np.any(valid_idx):
            print("警告: 没有有效数据")
            return None
        
        output_i_cleaned = output_i[valid_idx]
        time_cleaned = time_array[valid_idx]
        print(f"[DEBUG] len(output_i_cleaned): {len(output_i_cleaned)}, len(time_cleaned): {len(time_cleaned)}")
        
        if len(output_i_cleaned) < 2:
            print("警告: 有效数据点太少")
            return None
        
       
        output_power = output_i_cleaned**2 * r_load
        print(f"[DEBUG] output_power sample (first 5): {output_power[:5]}")
        
      
        output_energy = np.trapz(output_power, x=time_cleaned)
        print(f"[DEBUG] output_energy: {output_energy:.6f} J")
        
     
        input_duration = len(time_cleaned) * time_once
        print(f"[DEBUG] input_duration: {input_duration:.6f} s")
        input_energy = power_input * input_duration
        print(f"[DEBUG] input_energy (power_input * input_duration): {input_energy:.6f} J")
        
   
        efficiency = output_energy / input_energy if input_energy > 0 else 0.0
        print(f"[DEBUG] Calculated efficiency: {efficiency:.6f} ({efficiency*100:.2f}%)")
        
     
        avg_output_power = np.mean(output_power)
        max_output_power = np.max(output_power)
        
        return {
            "efficiency": efficiency,
            "avg_output_power": avg_output_power,
            "max_output_power": max_output_power,
            "output_energy": output_energy,
            "input_energy": input_energy,
            "duration": input_duration,
            "plot_data": {
                "time": time_cleaned,
                "current": output_i_cleaned,
                "power": output_power
            }
        }
        
    except Exception as e:
        print(f"简化计算过程中发生错误: {e}")
        return None

def calculate_unified_efficiencies(zheng_file_path: str, fan_file_path: str,
                                  reference_v: float, initial_v: float, r_load: float,
                                  drive_v: float, power_input: float,
                                  sampling_freq: float = 87500.0,
                                  points_to_process_zheng: int | None = None,
                                  points_to_process_fan: int | None = None):
   
    results = {
        "verification": { 
            "zheng": {"efficiency": 0.0, "stats": {}, "plot_data": {"time": np.array([]), "current": np.array([]), "power": np.array([])}},
            "fan": {"efficiency": 0.0, "stats": {}, "plot_data": {"time": np.array([]), "current": np.array([]), "power": np.array([])}},
            "finished_efficiency": 0.0
        },
        "theoretical": {  
            "zheng": {"efficiency": 0.0, "stats": {}, "plot_data": {"time": np.array([]), "output_current": np.array([]), "input_current": np.array([]), "output_power": np.array([]), "input_power": np.array([])}},
            "fan": {"efficiency": 0.0, "stats": {}, "plot_data": {"time": np.array([]), "output_current": np.array([]), "input_current": np.array([]), "output_power": np.array([]), "input_power": np.array([])}},
            "finished_efficiency": 0.0
        },
        "comparison": { 
            "zheng_diff": 0.0,
            "fan_diff": 0.0,
            "finished_diff": 0.0
        }
    }

    try:
        time_once = 1.0 / sampling_freq 
        

        data_zheng_df_full = pd.read_csv(zheng_file_path, header=None, skiprows=1)
        if points_to_process_zheng is not None and points_to_process_zheng > 0:
            data_zheng_df = data_zheng_df_full.head(points_to_process_zheng)
        else:
            data_zheng_df = data_zheng_df_full
        
        if data_zheng_df.empty:
            print(f"警告: 正接数据文件 '{zheng_file_path}' 为空或截取后为空。")
            return None
        else:
    
            channel_names = ["AIN1", "AIN2", "AIN3", "AIN4", "AIN5", "AIN6", "AIN7", "AIN8"]
            for i, channel in enumerate(channel_names):
                if data_zheng_df.shape[1] > i+1:
                    results["verification"]["zheng"]["stats"][channel] = _calculate_column_stats(data_zheng_df.iloc[:, i+1])
                    results["theoretical"]["zheng"]["stats"][channel] = _calculate_column_stats(data_zheng_df.iloc[:, i+1])
                else:
                    results["verification"]["zheng"]["stats"][channel] = {"max": np.nan, "min": np.nan, "avg": np.nan}
                    results["theoretical"]["zheng"]["stats"][channel] = {"max": np.nan, "min": np.nan, "avg": np.nan}

      
            t_zheng_original = pd.to_numeric(data_zheng_df.iloc[:, 0], errors='coerce').to_numpy() if data_zheng_df.shape[1] > 0 else np.array([])
            time_zheng = np.arange(len(t_zheng_original)) * time_once

   
            if data_zheng_df.shape[1] > 2:
          
                output_v_verification = pd.to_numeric(data_zheng_df.iloc[:, 2], errors='coerce').to_numpy()
                output_i_verification = (output_v_verification - initial_v) / reference_v
                
            
                valid_idx_ver = ~np.isnan(output_i_verification) & ~np.isnan(time_zheng)
                
                if np.any(valid_idx_ver):
                    output_i_ver_cleaned = output_i_verification[valid_idx_ver]
                    time_ver_cleaned = time_zheng[valid_idx_ver]
                    
                    if len(output_i_ver_cleaned) >= 2:
                      
                        output_power_ver = output_i_ver_cleaned**2 * r_load
                        
                        results["verification"]["zheng"]["plot_data"]["time"] = time_ver_cleaned
                        results["verification"]["zheng"]["plot_data"]["current"] = output_i_ver_cleaned
                        results["verification"]["zheng"]["plot_data"]["power"] = output_power_ver
                        
                  
                        output_energy = np.trapz(output_power_ver, x=time_ver_cleaned)
                        input_duration = len(time_ver_cleaned) * time_once
                        input_energy = power_input * input_duration
                        
                        if input_energy > 0:
                            results["verification"]["zheng"]["efficiency"] = output_energy / input_energy

       
            if data_zheng_df.shape[1] > 7:
                output_v_theoretical = pd.to_numeric(data_zheng_df.iloc[:, 6], errors='coerce').to_numpy()
                output_i_theoretical = (output_v_theoretical - initial_v) / reference_v
                
                input_v_theoretical = pd.to_numeric(data_zheng_df.iloc[:, 7], errors='coerce').to_numpy()
                input_i_theoretical = (input_v_theoretical - initial_v) / reference_v
                
                valid_idx_theo = ~np.isnan(output_i_theoretical) & ~np.isnan(input_i_theoretical) & ~np.isnan(time_zheng)
                
                if np.any(valid_idx_theo):
                    output_i_theo_cleaned = output_i_theoretical[valid_idx_theo]
                    input_i_theo_cleaned = input_i_theoretical[valid_idx_theo]
                    time_theo_cleaned = time_zheng[valid_idx_theo]
                    
                    if len(output_i_theo_cleaned) >= 2:
                        output_power_theo = output_i_theo_cleaned**2 * r_load
                        input_power_theo = drive_v * input_i_theo_cleaned
                        
                        results["theoretical"]["zheng"]["plot_data"]["time"] = time_theo_cleaned
                        results["theoretical"]["zheng"]["plot_data"]["output_current"] = output_i_theo_cleaned
                        results["theoretical"]["zheng"]["plot_data"]["input_current"] = input_i_theo_cleaned
                        results["theoretical"]["zheng"]["plot_data"]["output_power"] = output_power_theo
                        results["theoretical"]["zheng"]["plot_data"]["input_power"] = input_power_theo
                        
                        numerator = np.trapz(output_power_theo, x=time_theo_cleaned)
                        denominator = np.trapz(input_power_theo, x=time_theo_cleaned)
                        
                        if denominator > 0:
                            results["theoretical"]["zheng"]["efficiency"] = numerator / denominator

       
        data_fan_df_full = pd.read_csv(fan_file_path, header=None, skiprows=1)
        if points_to_process_fan is not None and points_to_process_fan > 0:
            data_fan_df = data_fan_df_full.head(points_to_process_fan)
        else:
            data_fan_df = data_fan_df_full

        if not data_fan_df.empty:
           
            for i, channel in enumerate(channel_names):
                if data_fan_df.shape[1] > i+1:
                    results["verification"]["fan"]["stats"][channel] = _calculate_column_stats(data_fan_df.iloc[:, i+1])
                    results["theoretical"]["fan"]["stats"][channel] = _calculate_column_stats(data_fan_df.iloc[:, i+1])
                else:
                    results["verification"]["fan"]["stats"][channel] = {"max": np.nan, "min": np.nan, "avg": np.nan}
                    results["theoretical"]["fan"]["stats"][channel] = {"max": np.nan, "min": np.nan, "avg": np.nan}

            t_fan_original = pd.to_numeric(data_fan_df.iloc[:, 0], errors='coerce').to_numpy() if data_fan_df.shape[1] > 0 else np.array([])
            time_fan = np.arange(len(t_fan_original)) * time_once

            if data_fan_df.shape[1] > 2:
                output_v_verification_fan = pd.to_numeric(data_fan_df.iloc[:, 2], errors='coerce').to_numpy()
                output_i_verification_fan = (output_v_verification_fan - initial_v) / reference_v
                
                valid_idx_ver_fan = ~np.isnan(output_i_verification_fan) & ~np.isnan(time_fan)
                
                if np.any(valid_idx_ver_fan):
                    output_i_ver_fan_cleaned = output_i_verification_fan[valid_idx_ver_fan]
                    time_ver_fan_cleaned = time_fan[valid_idx_ver_fan]
                    
                    if len(output_i_ver_fan_cleaned) >= 2:
                        output_power_ver_fan = output_i_ver_fan_cleaned**2 * r_load
                        
                        results["verification"]["fan"]["plot_data"]["time"] = time_ver_fan_cleaned
                        results["verification"]["fan"]["plot_data"]["current"] = output_i_ver_fan_cleaned
                        results["verification"]["fan"]["plot_data"]["power"] = output_power_ver_fan
                        
                        output_energy_fan = np.trapz(output_power_ver_fan, x=time_ver_fan_cleaned)
                        input_duration_fan = len(time_ver_fan_cleaned) * time_once
                        input_energy_fan = power_input * input_duration_fan
                        
                        if input_energy_fan > 0:
                            results["verification"]["fan"]["efficiency"] = output_energy_fan / input_energy_fan

            if data_fan_df.shape[1] > 7:
                output_v_theoretical_fan = pd.to_numeric(data_fan_df.iloc[:, 6], errors='coerce').to_numpy()
                output_i_theoretical_fan = (output_v_theoretical_fan - initial_v) / reference_v
                
                input_v_theoretical_fan = pd.to_numeric(data_fan_df.iloc[:, 7], errors='coerce').to_numpy()
                input_i_theoretical_fan = (input_v_theoretical_fan - initial_v) / reference_v
                
                valid_idx_theo_fan = ~np.isnan(output_i_theoretical_fan) & ~np.isnan(input_i_theoretical_fan) & ~np.isnan(time_fan)
                
                if np.any(valid_idx_theo_fan):
                    output_i_theo_fan_cleaned = output_i_theoretical_fan[valid_idx_theo_fan]
                    input_i_theo_fan_cleaned = input_i_theoretical_fan[valid_idx_theo_fan]
                    time_theo_fan_cleaned = time_fan[valid_idx_theo_fan]
                    
                    if len(output_i_theo_fan_cleaned) >= 2:
                        output_power_theo_fan = output_i_theo_fan_cleaned**2 * r_load
                        input_power_theo_fan = drive_v * input_i_theo_fan_cleaned
                        
                        results["theoretical"]["fan"]["plot_data"]["time"] = time_theo_fan_cleaned
                        results["theoretical"]["fan"]["plot_data"]["output_current"] = output_i_theo_fan_cleaned
                        results["theoretical"]["fan"]["plot_data"]["input_current"] = input_i_theo_fan_cleaned
                        results["theoretical"]["fan"]["plot_data"]["output_power"] = output_power_theo_fan
                        results["theoretical"]["fan"]["plot_data"]["input_power"] = input_power_theo_fan
                        
                        numerator_fan = np.trapz(output_power_theo_fan, x=time_theo_fan_cleaned)
                        denominator_fan = np.trapz(input_power_theo_fan, x=time_theo_fan_cleaned)
                        
                        if denominator_fan > 0:
                            results["theoretical"]["fan"]["efficiency"] = numerator_fan / denominator_fan

       
        ver_zheng_eff = max(0, results["verification"]["zheng"]["efficiency"])
        ver_fan_eff = max(0, results["verification"]["fan"]["efficiency"])
        results["verification"]["finished_efficiency"] = (ver_zheng_eff * ver_fan_eff) ** 0.5

        theo_zheng_raw = results["theoretical"]["zheng"]["efficiency"]
        theo_fan_raw = results["theoretical"]["fan"]["efficiency"]
        
        theo_zheng_processed = (max(0, theo_zheng_raw)) ** 0.5 * 100
        theo_fan_processed = (max(0, theo_fan_raw)) ** 0.5 * 100
        
        if theo_zheng_processed >= 0 and theo_fan_processed >= 0:
            results["theoretical"]["finished_efficiency"] = ((theo_zheng_processed * theo_fan_processed) / 10000) ** 0.5
        else:
            results["theoretical"]["finished_efficiency"] = 0.0
            
        results["theoretical"]["zheng"]["efficiency"] = theo_zheng_processed / 100
        results["theoretical"]["fan"]["efficiency"] = theo_fan_processed / 100

        results["comparison"]["zheng_diff"] = abs(results["theoretical"]["zheng"]["efficiency"] - results["verification"]["zheng"]["efficiency"])
        results["comparison"]["fan_diff"] = abs(results["theoretical"]["fan"]["efficiency"] - results["verification"]["fan"]["efficiency"])
        results["comparison"]["finished_diff"] = abs(results["theoretical"]["finished_efficiency"] - results["verification"]["finished_efficiency"])

        return results

    except FileNotFoundError as e:
        print(f"错误: CSV文件未找到。 {e}")
        return None
    except pd.errors.EmptyDataError as e:
        print(f"错误: CSV文件为空或解析后无数据。 {e}") 
        return None
    except pd.errors.ParserError as e:
        print(f"错误: 解析CSV文件时出错。请检查文件格式。 {e}")
        return None
    except Exception as e:
        print(f"统一计算过程中发生未预料的错误: {e}")
        return None 



class ExperimentConfig:
    """实验配置类，用于管理不同探究因素的参数设置"""
    
    def __init__(self, is_factor_exploration_mode: bool = False):
        self.exploration_type = None 
        self.fixed_params = {}
        self.variable_params = []
        self.common_params = {
            'reference_v': 1.0,     
            'initial_v': 0.0,       
            'sampling_freq': 87500.0  
        }
        self.is_factor_exploration_mode = is_factor_exploration_mode 
    
    def configure_voltage_exploration(self, voltage_levels, r_load_fixed=None):
     
        self.exploration_type = 'voltage'
        if r_load_fixed is not None:
            self.fixed_params['r_load'] = r_load_fixed
        # else: r_load will be taken from common_params
        self.variable_params = voltage_levels
    
    def configure_resistance_exploration(self, resistance_power_levels, drive_v_fixed=12.0):
       
        self.exploration_type = 'resistance'
        self.fixed_params = {
            'drive_v': drive_v_fixed,
            # 'power_input': power_input_fixed #不再固定power_input
        }
        # self.variable_params = [{'r_load': r} for r in resistance_levels]
        self.variable_params = resistance_power_levels
    
    def configure_magnetic_distance_exploration(self, distance_levels, drive_v_fixed=12.0, r_load_fixed=None):
       
        self.exploration_type = 'magnetic_distance'
        self.fixed_params = {
            'drive_v': drive_v_fixed,
        }
        if r_load_fixed is not None:
            self.fixed_params['r_load'] = r_load_fixed
   
        self.variable_params = []
        for distance in distance_levels:
           
            power_input = 10.0 
            self.variable_params.append({
                'magnetic_distance': distance,
                'power_input': power_input
            })
    
    def get_experiment_params(self, index):
     
        if index >= len(self.variable_params):
            raise IndexError(f"实验组索引 {index} 超出范围")
        
        params = self.common_params.copy()
        params.update(self.fixed_params)
        params.update(self.variable_params[index])
        
        return params
    
    def save_config(self, filepath):
      
        config_data = {
            'exploration_type': self.exploration_type,
            'fixed_params': self.fixed_params,
            'variable_params': self.variable_params,
            'common_params': self.common_params,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=2)
    
    @classmethod
    def load_config(cls, filepath):
       
        with open(filepath, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        
        config = cls()
        config.exploration_type = config_data['exploration_type']
        config.fixed_params = config_data['fixed_params']
        config.variable_params = config_data['variable_params']
        config.common_params = config_data['common_params']
        
        return config


class BatchExperimentAnalyzer:
  
    
    def __init__(self, config: ExperimentConfig):
        self.config = config
        self.results = []
    
    def run_batch_experiments(self, file_pattern_or_zheng: str, 
                            fan_file_pattern: str | None = None,
                            points_to_process: int | None = None): 
       
        self.results = []
        
        for i in range(len(self.config.variable_params)):
            params_from_config = self.config.get_experiment_params(i)
            
            current_file_path = file_pattern_or_zheng.format(index=i+1)
            
            if self.config.is_factor_exploration_mode:
               
                if not os.path.exists(current_file_path):
                    print(f"警告: 第 {i+1} 组因素探究文件 '{current_file_path}' 未找到，跳过")
                    continue
                print(f"\n运行第 {i+1} 组因素探究: {self.config.exploration_type}")
                print(f"文件: {current_file_path}")
                print(f"参数: {params_from_config}")
                
               
                result = calculate_unified_efficiencies(
                    zheng_file_path=current_file_path,
                    fan_file_path=current_file_path, 
                    reference_v=params_from_config['reference_v'],
                    initial_v=params_from_config['initial_v'],
                    r_load=params_from_config['r_load'],
                    drive_v=params_from_config.get('drive_v', 0),
                    power_input=params_from_config['power_input'],
                    sampling_freq=params_from_config['sampling_freq'],
                    points_to_process_zheng=points_to_process,
                    points_to_process_fan=points_to_process 
                )
                if result:
                   
                    factor_efficiency = result["verification"]["finished_efficiency"]
                    simplified_result = {
                        'experiment_params': params_from_config,
                        'experiment_index': i + 1,
                        'factor_exploration_mode': True,
                        'efficiency': factor_efficiency,
                      
                        'plot_data': result["verification"]["zheng"]["plot_data"], 
                        'avg_output_power': np.mean(result["verification"]["zheng"]["plot_data"]["power"]) if len(result["verification"]["zheng"]["plot_data"]["power"]) > 0 else 0,
                        'max_output_power': np.max(result["verification"]["zheng"]["plot_data"]["power"]) if len(result["verification"]["zheng"]["plot_data"]["power"]) > 0 else 0,
                    }
                    self.results.append(simplified_result)
            else:
               
                if fan_file_pattern is None:
                    print("错误: 双机标定模式需要提供反接文件模式")
                    return
                
                zheng_file = current_file_path
                fan_file = fan_file_pattern.format(index=i+1)
                
                if not os.path.exists(zheng_file) or not os.path.exists(fan_file):
                    print(f"警告: 第 {i+1} 组双机标定文件不完整 (Z: {zheng_file}, F: {fan_file})，跳过")
                    continue
                
                print(f"\n运行第 {i+1} 组双机标定实验...")
                print(f"参数: {params_from_config}")
                
                result = calculate_unified_efficiencies(
                    zheng_file_path=zheng_file,
                    fan_file_path=fan_file,
                    reference_v=params_from_config['reference_v'],
                    initial_v=params_from_config['initial_v'],
                    r_load=params_from_config['r_load'],
                    drive_v=params_from_config.get('drive_v', 0),
                    power_input=params_from_config['power_input'],
                    sampling_freq=params_from_config['sampling_freq'],
                    points_to_process_zheng=points_to_process,
                    points_to_process_fan=points_to_process
                )
                if result:
                    result['experiment_params'] = params_from_config
                    result['experiment_index'] = i + 1
                    self.results.append(result)
    
    def generate_comparison_table(self):
       
        if not self.results:
            print("错误: 没有可用的实验结果")
            return None
        
       
        table_data = []
        
        for i, result in enumerate(self.results):
            params = result['experiment_params']
            row = {
                '实验组': result['experiment_index']
            }
            
          
            if self.config.exploration_type == 'voltage':
                row['输入电压(V)'] = params['drive_v']
                row['输入功率(W)'] = params['power_input']
            elif self.config.exploration_type == 'resistance':
                row['负载电阻(Ω)'] = params['r_load']
                row['输入功率(W)'] = params['power_input']
            else: 
              
                var_params = self.config.variable_params[i]
                if 'magnetic_distance' in var_params:
                    row['磁场距离(mm)'] = var_params['magnetic_distance']
                row['输入功率(W)'] = params['power_input']
            
          
            if 'simple_mode' in result and result['simple_mode']:
               
                row['效率(%)'] = f"{result['efficiency']*100:.2f}"
                row['平均输出功率(W)'] = f"{result['avg_output_power']:.2f}"
                row['最大输出功率(W)'] = f"{result['max_output_power']:.2f}"
            else:
               
                row['验证实验-正接效率'] = f"{result['verification']['zheng']['efficiency']:.4f}"
                row['验证实验-反接效率'] = f"{result['verification']['fan']['efficiency']:.4f}"
                row['验证实验-综合效率'] = f"{result['verification']['finished_efficiency']:.4f}"
                
                row['理论实验-正接效率'] = f"{result['theoretical']['zheng']['efficiency']:.4f}"
                row['理论实验-反接效率'] = f"{result['theoretical']['fan']['efficiency']:.4f}"
                row['理论实验-综合效率'] = f"{result['theoretical']['finished_efficiency']:.4f}"
                
                row['正接效率差值'] = f"{result['comparison']['zheng_diff']:.4f}"
                row['反接效率差值'] = f"{result['comparison']['fan_diff']:.4f}"
                row['综合效率差值'] = f"{result['comparison']['finished_diff']:.4f}"
            
            table_data.append(row)
        
     
        df = pd.DataFrame(table_data)
        
      
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = f'实验结果对比_{self.config.exploration_type}_{timestamp}.xlsx'
        
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='实验结果对比', index=False)
            
           
            config_df = pd.DataFrame([
                ['探究类型', self.config.exploration_type],
                ['固定参数', str(self.config.fixed_params)],
                ['通用参数', str(self.config.common_params)],
                ['实验时间', timestamp]
            ], columns=['参数名', '参数值'])
            config_df.to_excel(writer, sheet_name='实验配置', index=False)
        
        print(f"\n实验结果已保存到: {excel_filename}")
        
        return df
    
    def plot_efficiency_curves(self):
      
        if not self.results:
            print("错误: 没有可用的实验结果")
            return
        
      
        is_simple_mode = any('simple_mode' in result and result['simple_mode'] for result in self.results)
        
        if is_simple_mode:
        
            x_values = []
            efficiencies = []
            
            for i, result in enumerate(self.results):
                params = result['experiment_params']
                
                if self.config.exploration_type == 'voltage':
                    x_values.append(params['drive_v'])
                    x_label = '输入电压 (V)'
                    title_prefix = '电压'
                elif self.config.exploration_type == 'resistance':
                    x_values.append(params['r_load'])
                    x_label = '负载电阻 (Ω)'
                    title_prefix = '负载电阻'
                else: 
                 
                    var_params = self.config.variable_params[i]
                    if 'magnetic_distance' in var_params:
                        x_values.append(var_params['magnetic_distance'])
                    else:
                        x_values.append(0.0)
                    x_label = '磁场距离 (mm)'
                    title_prefix = '磁场距离'
                
                efficiencies.append(result.get('efficiency', 0) * 100)
            
          
            fig, ax = plt.subplots(1, 1, figsize=(10, 6))
            
            ax.plot(x_values, efficiencies, 'o-', label='效率', markersize=10, linewidth=2, color='blue')
            
           
            if efficiencies:
                max_idx = np.argmax(efficiencies)
                ax.scatter([x_values[max_idx]], [efficiencies[max_idx]], 
                          s=200, c='red', marker='*', 
                          label=f'最高效率: {efficiencies[max_idx]:.2f}%')
            
            ax.set_xlabel(x_label)
            ax.set_ylabel('效率 (%)')
            ax.set_title(f'效率随{title_prefix}变化曲线')
            ax.legend()
            ax.grid(True, alpha=0.3)
            
          
            for x, y in zip(x_values, efficiencies):
                ax.annotate(f'{y:.1f}%', (x, y), textcoords="offset points", 
                           xytext=(0,10), ha='center')
        else:
           
            x_values = []
            ver_zheng_eff = []
            ver_fan_eff = []
            ver_finished_eff = []
            theo_zheng_eff = []
            theo_fan_eff = []
            theo_finished_eff = []
            
            for i, result in enumerate(self.results):
                params = result['experiment_params']
                
                if self.config.exploration_type == 'voltage':
                    x_values.append(params['drive_v'])
                    x_label = '输入电压 (V)'
                elif self.config.exploration_type == 'resistance':
                    x_values.append(params['r_load'])
                    x_label = '负载电阻 (Ω)'
                else: 
                   
                    var_params = self.config.variable_params[i]
                    if 'magnetic_distance' in var_params:
                        x_values.append(var_params['magnetic_distance'])
                    else:
                        x_values.append(0.0)
                    x_label = '磁场距离 (mm)'
                
                ver_zheng_eff.append(result['verification']['zheng']['efficiency'])
                ver_fan_eff.append(result['verification']['fan']['efficiency'])
                ver_finished_eff.append(result['verification']['finished_efficiency'])
                
                theo_zheng_eff.append(result['theoretical']['zheng']['efficiency'])
                theo_fan_eff.append(result['theoretical']['fan']['efficiency'])
                theo_finished_eff.append(result['theoretical']['finished_efficiency'])
            
        
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
            
           
            ax1.plot(x_values, ver_zheng_eff, 'o-', label='正接效率', markersize=8)
            ax1.plot(x_values, ver_fan_eff, 's-', label='反接效率', markersize=8)
            ax1.plot(x_values, ver_finished_eff, '^-', label='综合效率', markersize=8)
            ax1.set_xlabel(x_label)
            ax1.set_ylabel('效率')
            ax1.set_title('验证实验效率对比')
            ax1.legend()
            ax1.grid(True, alpha=0.3)
            
          
            ax2.plot(x_values, theo_zheng_eff, 'o-', label='正接效率', markersize=8)
            ax2.plot(x_values, theo_fan_eff, 's-', label='反接效率', markersize=8)
            ax2.plot(x_values, theo_finished_eff, '^-', label='综合效率', markersize=8)
            ax2.set_xlabel(x_label)
            ax2.set_ylabel('效率')
            ax2.set_title('理论实验效率对比')
            ax2.legend()
            ax2.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
       
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        fig_filename = f'效率曲线_{self.config.exploration_type}_{timestamp}.png'
        plt.savefig(fig_filename, dpi=300, bbox_inches='tight')
        plt.show()
        
        print(f"\n效率曲线图已保存到: {fig_filename}")



def run_voltage_exploration_example():
    """运行输入电压探究实验示例"""
    print("=== 探究输入电压对效率的影响 ===")
    
    config = ExperimentConfig()
    voltage_levels = [
        {'drive_v': 6.0, 'power_input': 5.0},
        {'drive_v': 9.0, 'power_input': 8.0},
        {'drive_v': 12.0, 'power_input': 12.0},
        {'drive_v': 15.0, 'power_input': 18.0},
        {'drive_v': 18.0, 'power_input': 25.0}
    ]
    config.configure_voltage_exploration(voltage_levels, r_load_fixed=10.0)
    config.save_config('voltage_exploration_config.json')
    
    analyzer = BatchExperimentAnalyzer(config)
   


def run_resistance_exploration_example():
   
    print("=== 探究负载电阻对效率的影响 ===")
    
    config = ExperimentConfig()
    resistance_levels = [5.0, 10.0, 15.0, 20.0, 25.0, 30.0]
    config.configure_resistance_exploration(
        resistance_levels, 
        drive_v_fixed=12.0, 
        power_input_fixed=10.0
    )
    config.save_config('resistance_exploration_config.json')
    
    analyzer = BatchExperimentAnalyzer(config)
   


if __name__ == "__main__":

    print("提供单次实验和批量实验两种模式")
    print("1. 单次实验: 直接调用 calculate_unified_efficiencies 函数")
    print("2. 批量实验: 使用 ExperimentConfig 和 BatchExperimentAnalyzer 类")
    print("\n详见 run_voltage_exploration_example() 和 run_resistance_exploration_example() 函数") 