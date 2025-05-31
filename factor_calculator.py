import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional

def calculate_single_efficiency(csv_file_path: str,
                              reference_v: float,
                              initial_v: float,
                              r_load: float,
                              power_input: float,
                              sampling_freq: float = 87500.0,
                              points_to_process: Optional[int] = None) -> Dict:
    try:
        data_df = pd.read_csv(csv_file_path, header=None, skiprows=1)
        
        if points_to_process is not None and points_to_process > 0:
            data_df = data_df.head(points_to_process)
        
        if data_df.empty or data_df.shape[1] < 3:
            raise ValueError("数据文件为空或列数不足")
        
        time_once = 1.0 / sampling_freq
        time_array = np.arange(len(data_df)) * time_once
        
        output_v = pd.to_numeric(data_df.iloc[:, 2], errors='coerce').to_numpy()
        output_i = (output_v - initial_v) / reference_v
        
        valid_idx = ~np.isnan(output_i)
        if not np.any(valid_idx):
            raise ValueError("没有有效的电流数据")
        
        output_i_cleaned = output_i[valid_idx]
        time_cleaned = time_array[valid_idx]
        
        output_power = output_i_cleaned**2 * r_load
        
        output_energy = np.trapz(output_power, x=time_cleaned)
        input_energy = power_input * (len(time_cleaned) * time_once)
        
        efficiency = output_energy / input_energy if input_energy > 0 else 0.0
        
        return {
            "efficiency": efficiency,
            "avg_output_power": np.mean(output_power),
            "max_output_power": np.max(output_power),
            "avg_output_current": np.mean(output_i_cleaned),
            "duration": len(time_cleaned) * time_once,
            "plot_data": {
                "time": time_cleaned,
                "current": output_i_cleaned,
                "power": output_power
            }
        }
    
    except Exception as e:
        print(f"计算效率时出错: {e}")
        return None

def calculate_factor_experiment(experiment_data: List[Dict]) -> Dict:
    results = {
        "factor_values": [],
        "efficiencies": [],
        "labels": [],
        "avg_powers": [],
        "trend_analysis": {}
    }
    
    for exp in experiment_data:
        factor_value = exp.get("factor_value")
        file_path = exp.get("file_path")
        label = exp.get("label", f"Factor={factor_value}")
        
        calc_params = exp.get("params", {})
        
        result = calculate_single_efficiency(
            csv_file_path=file_path,
            **calc_params
        )
        
        if result:
            results["factor_values"].append(factor_value)
            results["efficiencies"].append(result["efficiency"])
            results["labels"].append(label)
            results["avg_powers"].append(result["avg_output_power"])
    
    if len(results["factor_values"]) >= 2:
        factor_array = np.array(results["factor_values"])
        eff_array = np.array(results["efficiencies"])
        
        if len(factor_array) >= 2:
            coeffs = np.polyfit(factor_array, eff_array, 1)
            results["trend_analysis"]["linear_slope"] = coeffs[0]
            results["trend_analysis"]["linear_intercept"] = coeffs[1]
            
        max_idx = np.argmax(eff_array)
        results["trend_analysis"]["optimal_factor"] = factor_array[max_idx]
        results["trend_analysis"]["optimal_efficiency"] = eff_array[max_idx]
        
        results["trend_analysis"]["efficiency_range"] = np.ptp(eff_array)
        results["trend_analysis"]["relative_change"] = np.ptp(eff_array) / np.mean(eff_array) * 100
    
    return results

def compare_dual_motor_efficiencies(zheng_file: str, fan_file: str, 
                                   reference_v: float, initial_v: float,
                                   r_load: float, power_input: float,
                                   sampling_freq: float = 87500.0) -> float:
    zheng_result = calculate_single_efficiency(
        zheng_file, reference_v, initial_v, r_load, 
        power_input, sampling_freq
    )
    
    fan_result = calculate_single_efficiency(
        fan_file, reference_v, initial_v, r_load,
        power_input, sampling_freq
    )
    
    if zheng_result and fan_result:
        zheng_eff = max(0, zheng_result["efficiency"])
        fan_eff = max(0, fan_result["efficiency"])
        combined_efficiency = (zheng_eff * fan_eff) ** 0.5
        return combined_efficiency
    
    return 0.0 