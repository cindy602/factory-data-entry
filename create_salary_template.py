#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
制造工厂工资计算模板生成器
根据生产数据自动生成工资计算Excel模板
"""

import pandas as pd
from datetime import datetime, timedelta
import os

def create_salary_template():
    """创建工资计算Excel模板"""
    
    # 获取当前月份
    current_month = datetime.now().strftime('%Y年%m月')
    template_file = f'工资计算模板_{current_month}.xlsx'
    
    # 创建工资计算模板
    with pd.ExcelWriter(template_file, engine='openpyxl') as writer:
        # 1. 员工工资计算表
        salary_data = {
            '员工姓名': ['员工1', '员工2', '员工3', '员工4', '员工5'],
            '员工编号': ['emp001', 'emp002', 'emp003', 'emp004', 'emp005'],
            '计件单价(元/件)': [1.5, 1.5, 1.5, 1.5, 1.5],
            '总产量(件)': [0, 0, 0, 0, 0],
            '计件工资(元)': [0, 0, 0, 0, 0],
            '质量奖金(元)': [0, 0, 0, 0, 0],
            '全勤奖金(元)': [200, 200, 200, 200, 200],
            '其他补贴(元)': [0, 0, 0, 0, 0],
            '应发工资(元)': [0, 0, 0, 0, 0],
            '扣款(元)': [0, 0, 0, 0, 0],
            '实发工资(元)': [0, 0, 0, 0, 0],
            '备注': ['', '', '', '', '']
        }
        
        df_salary = pd.DataFrame(salary_data)
        df_salary.to_excel(writer, sheet_name='工资计算', index=False)
        
        # 2. 产量统计表
        production_data = {
            '日期': [],
            '员工姓名': [],
            '款号': [],
            '颜色': [],
            '数量': [],
            '质量状态': [],
            '工序': [],
            '计件工资': []
        }
        
        df_production = pd.DataFrame(production_data)
        df_production.to_excel(writer, sheet_name='产量明细', index=False)
        
        # 3. 质量统计表
        quality_data = {
            '员工姓名': ['员工1', '员工2', '员工3', '员工4', '员工5'],
            '合格数量': [0, 0, 0, 0, 0],
            '不合格数量': [0, 0, 0, 0, 0],
            '返工数量': [0, 0, 0, 0, 0],
            '合格率': ['0%', '0%', '0%', '0%', '0%'],
            '质量奖金': [0, 0, 0, 0, 0]
        }
        
        df_quality = pd.DataFrame(quality_data)
        df_quality.to_excel(writer, sheet_name='质量统计', index=False)
        
        # 4. 工资汇总表
        summary_data = {
            '项目': ['总员工数', '总产量', '总工资支出', '平均工资', '最高工资', '最低工资'],
            '数值': [5, 0, 0, 0, 0, 0],
            '单位': ['人', '件', '元', '元', '元', '元']
        }
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='工资汇总', index=False)
    
    print(f"✓ 工资计算模板已创建: {template_file}")
    print("模板包含以下工作表:")
    print("1. 工资计算 - 员工工资明细计算")
    print("2. 产量明细 - 生产数据明细")
    print("3. 质量统计 - 质量数据统计")
    print("4. 工资汇总 - 工资支出汇总")
    
    # 添加公式说明
    print("\n使用说明:")
    print("1. 将生产数据从 data_records.xlsx 复制到 '产量明细' 表")
    print("2. 在 '工资计算' 表中:")
    print("   - 计件工资 = 总产量 × 计件单价")
    print("   - 应发工资 = 计件工资 + 质量奖金 + 全勤奖金 + 其他补贴")
    print("   - 实发工资 = 应发工资 - 扣款")
    print("3. 质量奖金根据合格率计算（合格率>95%: 100元，>90%: 50元）")
    
    return template_file

def update_salary_from_production(salary_template, production_data):
    """根据生产数据更新工资模板"""
    try:
        # 读取生产数据
        if os.path.exists(production_data):
            df_production = pd.read_excel(production_data)
            
            # 读取工资模板
            df_salary = pd.read_excel(salary_template, sheet_name='工资计算')
            
            # 按员工统计产量
            if '人员' in df_production.columns and '数量' in df_production.columns:
                production_by_person = df_production.groupby('人员')['数量'].sum()
                
                # 更新工资表中的总产量
                for idx, row in df_salary.iterrows():
                    employee_name = row['员工姓名']
                    if employee_name in production_by_person:
                        df_salary.at[idx, '总产量(件)'] = production_by_person[employee_name]
                        df_salary.at[idx, '计件工资(元)'] = production_by_person[employee_name] * row['计件单价(元/件)']
            
            # 保存更新后的工资表
            with pd.ExcelWriter(salary_template, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_salary.to_excel(writer, sheet_name='工资计算', index=False)
            
            print(f"✓ 已根据生产数据更新工资模板")
            
    except Exception as e:
        print(f"✗ 更新工资模板失败: {e}")

if __name__ == '__main__':
    print("=" * 50)
    print("制造工厂工资计算模板生成器")
    print("=" * 50)
    
    # 创建工资模板
    template_file = create_salary_template()
    
    # 询问是否要基于现有生产数据更新模板
    production_file = 'data_records.xlsx'
    if os.path.exists(production_file):
        response = input("\n是否基于现有生产数据更新工资模板? (y/n): ")
        if response.lower() == 'y':
            update_salary_from_production(template_file, production_file)
    
    print("\n" + "=" * 50)
    print("操作完成!")
    print(f"工资模板文件: {template_file}")
    print("下一步:")
    print("1. 打开Excel文件查看和编辑工资模板")
    print("2. 根据实际需求调整计件单价和奖金标准")
    print("3. 每月导出生产数据后使用此模板计算工资")
    print("=" * 50)