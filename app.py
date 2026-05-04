#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
制造工厂数据录入系统 v3.3
改进：
1. 增加尺码字段
2. 员工管理（管理员可增删改员工登录名、密码）
3. 统计按月筛选
"""

import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template, session, redirect, url_for
from datetime import datetime
import json
import hashlib
import shutil
import tempfile

app = Flask(__name__)
app.secret_key = 'factory_data_entry_2024_v3'

# ========== 配置 ==========
FACTORIES = {
    'factory1': {'name': '一厂', 'file': 'data_factory1.xlsx'},
    'factory2': {'name': '二厂', 'file': 'data_factory2.xlsx'}
}

EMPLOYEES_FILE = 'employees.json'
PRICE_FILE = 'prices.json'
BACKUP_DIR = 'backups'

# ========== Excel列定义 ==========
EXCEL_COLUMNS = ['流水号', '日期', '款号', '颜色', '尺码', '数量', '床数',
                 '质量状态', '人员', '所属工厂', '备注', '提交时间']

# ========== 初始化 ==========
def init_excel():
    for factory_id, config in FACTORIES.items():
        file_path = config['file']
        if not os.path.exists(file_path):
            df = pd.DataFrame(columns=EXCEL_COLUMNS)
            df.to_excel(file_path, index=False)
            print(f"✓ 已创建 {config['name']} 数据文件: {file_path}")
        else:
            df = pd.read_excel(file_path)
            for col in EXCEL_COLUMNS:
                if col not in df.columns:
                    df[col] = ''
            df = df[EXCEL_COLUMNS]
            df.to_excel(file_path, index=False)
            print(f"✓ {config['name']} 数据文件已就绪: {file_path}")

def init_employees():
    if not os.path.exists(EMPLOYEES_FILE):
        employees = {
            "admin": {
                "password": hashlib.md5("admin123".encode()).hexdigest(),
                "name": "管理员",
                "role": "admin",
                "factory": "all"
            }
        }
        for i in range(1, 11):
            emp_id = f"f1_emp{i:02d}"
            employees[emp_id] = {
                "password": hashlib.md5(f"f1_emp{i:02d}123".encode()).hexdigest(),
                "name": f"一厂-员工{i}",
                "role": "employee",
                "factory": "factory1"
            }
        for i in range(1, 11):
            emp_id = f"f2_emp{i:02d}"
            employees[emp_id] = {
                "password": hashlib.md5(f"f2_emp{i:02d}123".encode()).hexdigest(),
                "name": f"二厂-员工{i}",
                "role": "employee",
                "factory": "factory2"
            }
        with open(EMPLOYEES_FILE, 'w', encoding='utf-8') as f:
            json.dump(employees, f, ensure_ascii=False, indent=2)
        print(f"✓ 已创建员工配置文件")
    else:
        print(f"✓ 员工配置文件已存在")

def init_prices():
    if not os.path.exists(PRICE_FILE):
        with open(PRICE_FILE, 'w', encoding='utf-8') as f:
            json.dump({}, f, ensure_ascii=False, indent=2)
        print(f"✓ 已创建单价配置文件: {PRICE_FILE}")
    else:
        print(f"✓ 单价配置文件已存在")

def init_backup():
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
        print(f"✓ 已创建备份目录: {BACKUP_DIR}")

# ========== 工具函数 ==========
def safe_int(value, default=0):
    try:
        if value is None or (isinstance(value, float) and (pd.isna(value) or value != value)):
            return default
        return int(float(str(value).strip()))
    except:
        return default

def safe_float(value, default=0.0):
    try:
        if value is None or (isinstance(value, float) and (pd.isna(value) or value != value)):
            return default
        return float(str(value).strip())
    except:
        return default

def safe_str(value, default=''):
    try:
        if value is None or (isinstance(value, float) and (pd.isna(value) or value != value)):
            return default
        return str(value).strip()
    except:
        return default

def clean_records_for_json(records):
    cleaned = []
    for r in records:
        clean = {}
        for k, v in r.items():
            if isinstance(v, float) and (pd.isna(v) or v != v):
                clean[k] = None
            else:
                clean[k] = v
        cleaned.append(clean)
    return cleaned

def get_factory_file(factory_id):
    if factory_id in FACTORIES:
        return FACTORIES[factory_id]['file']
    return None

def add_to_excel(data, factory_id):
    try:
        file_path = get_factory_file(factory_id)
        if not file_path:
            return False, f"无效的工厂ID: {factory_id}"
        
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
        else:
            df = pd.DataFrame(columns=EXCEL_COLUMNS)
        
        new_row = {
            '流水号': safe_str(data.get('serial_no', '')),
            '日期': safe_str(data.get('date', '')),
            '款号': safe_str(data.get('style_no', '')),
            '颜色': safe_str(data.get('color', '')),
            '尺码': safe_str(data.get('size', '')),
            '数量': safe_int(data.get('quantity', 0)),
            '床数': safe_str(data.get('bed_count', '')),
            '质量状态': safe_str(data.get('quality', '合格')),
            '人员': safe_str(data.get('person', '')),
            '所属工厂': FACTORIES[factory_id]['name'],
            '备注': safe_str(data.get('remark', '')),
            '提交时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(file_path, index=False)
        
        create_daily_backup(factory_id)
        
        return True, "数据保存成功"
    except Exception as e:
        return False, f"保存失败: {str(e)}"

def create_daily_backup(factory_id):
    try:
        file_path = get_factory_file(factory_id)
        if file_path and os.path.exists(file_path):
            today = datetime.now().strftime('%Y%m%d')
            backup_file = os.path.join(BACKUP_DIR, f'{factory_id}_backup_{today}.xlsx')
            if not os.path.exists(backup_file):
                shutil.copy2(file_path, backup_file)
    except Exception as e:
        print(f"✗ 创建备份失败: {e}")

# ========== 单价管理 ==========
def load_prices():
    try:
        if os.path.exists(PRICE_FILE):
            with open(PRICE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    except:
        return {}

def save_prices(prices):
    with open(PRICE_FILE, 'w', encoding='utf-8') as f:
        json.dump(prices, f, ensure_ascii=False, indent=2)

# ========== 用户认证 ==========
def load_employees():
    try:
        if os.path.exists(EMPLOYEES_FILE):
            with open(EMPLOYEES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    except:
        return {}

def save_employees(employees):
    with open(EMPLOYEES_FILE, 'w', encoding='utf-8') as f:
        json.dump(employees, f, ensure_ascii=False, indent=2)

def verify_login(username, password):
    employees = load_employees()
    if username in employees:
        hashed_password = hashlib.md5(password.encode()).hexdigest()
        if employees[username]['password'] == hashed_password:
            return employees[username]
    return None

# ========== 路由 ==========
@app.route('/')
def index():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    factory_name = ''
    if session.get('factory') and session['factory'] in FACTORIES:
        factory_name = FACTORIES[session['factory']]['name']
    
    return render_template('index.html', 
                         username=session.get('username'),
                         role=session.get('role'),
                         name=session.get('name'),
                         factory=session.get('factory'),
                         factory_name=factory_name)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html', factories=FACTORIES)
    
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()
    
    if not username or not password:
        return render_template('login.html', error='用户名和密码不能为空', factories=FACTORIES)
    
    user_info = verify_login(username, password)
    if user_info:
        session['username'] = username
        session['role'] = user_info['role']
        session['name'] = user_info['name']
        session['factory'] = user_info.get('factory', 'factory1')
        return redirect(url_for('index'))
    else:
        return render_template('login.html', error='用户名或密码错误', factories=FACTORIES)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ---- 数据提交 ----
@app.route('/api/submit', methods=['POST'])
def submit_data():
    try:
        if 'username' not in session:
            return jsonify({'success': False, 'message': '请先登录'}), 401
        
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': '未收到数据'}), 400
        
        required_fields = ['date', 'style_no', 'color', 'quantity']
        for field in required_fields:
            if field not in data or not str(data[field]).strip():
                return jsonify({'success': False, 'message': f'字段{field}不能为空'}), 400
        
        data['person'] = session.get('name', session.get('username'))
        factory_id = session.get('factory', 'factory1')
        
        success, message = add_to_excel(data, factory_id)
        if success:
            return jsonify({'success': True, 'message': '✅ 生产数据保存成功！'})
        else:
            return jsonify({'success': False, 'message': message}), 500
    except Exception as e:
        return jsonify({'success': False, 'message': f'服务器错误: {str(e)}'}), 500

# ---- 查询记录 ----
@app.route('/api/records')
def get_records():
    try:
        if 'username' not in session:
            return jsonify({'success': False, 'message': '请先登录'}), 401
        
        all_records = []
        prices = load_prices()
        
        if session.get('role') == 'admin':
            for factory_id, config in FACTORIES.items():
                file_path = config['file']
                if os.path.exists(file_path):
                    df = pd.read_excel(file_path)
                    if not df.empty:
                        records = df.to_dict('records')
                        for r in records:
                            style = safe_str(r.get('款号', ''))
                            price = prices.get(style, 0)
                            qty = safe_int(r.get('数量', 0))
                            r['单价'] = price
                            r['金额'] = round(price * qty, 2)
                        all_records.extend(records)
        else:
            factory_id = session.get('factory', 'factory1')
            file_path = get_factory_file(factory_id)
            if file_path and os.path.exists(file_path):
                df = pd.read_excel(file_path)
                current_user = session.get('name', session.get('username'))
                if '人员' in df.columns and not df.empty:
                    df = df[df['人员'] == current_user]
                if not df.empty:
                    records = df.to_dict('records')
                    for r in records:
                        style = safe_str(r.get('款号', ''))
                        price = prices.get(style, 0)
                        qty = safe_int(r.get('数量', 0))
                        r['单价'] = price
                        r['金额'] = round(price * qty, 2)
                    all_records = records
        
        all_records = clean_records_for_json(all_records)
        return jsonify({'success': True, 'records': all_records})
    except Exception as e:
        return jsonify({'success': False, 'message': f'读取数据失败: {str(e)}'}), 500

# ---- 统计（按月筛选） ----
@app.route('/api/statistics')
def get_statistics():
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        month = request.args.get('month', '')
        
        statistics = {
            'total_records': 0,
            'total_quantity': 0,
            'by_factory': {},
            'by_person': {},
            'by_style': {},
            'by_color': {},
            'by_quality': {},
            'monthly_salary': {},
            'style_summary': {},
            'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        prices = load_prices()
        
        for factory_id, config in FACTORIES.items():
            file_path = config['file']
            if os.path.exists(file_path):
                df = pd.read_excel(file_path)
                if df.empty:
                    continue
                
                # 按月筛选
                if month and '日期' in df.columns:
                    df = df[df['日期'].astype(str).str.startswith(month)]
                    if df.empty:
                        continue
                
                if '数量' in df.columns:
                    df['数量'] = pd.to_numeric(df['数量'], errors='coerce').fillna(0).astype(int)
                
                factory_name = config['name']
                factory_records = len(df)
                factory_quantity = int(df['数量'].sum()) if '数量' in df.columns else 0
                
                statistics['total_records'] += factory_records
                statistics['total_quantity'] += factory_quantity
                
                statistics['by_factory'][factory_name] = {
                    'records': factory_records,
                    'quantity': factory_quantity
                }
                
                if '人员' in df.columns and '数量' in df.columns and '款号' in df.columns:
                    for person, group in df.groupby('人员'):
                        total_qty = int(group['数量'].sum())
                        total_amount = 0
                        for _, row in group.iterrows():
                            style = safe_str(row.get('款号', ''))
                            qty = safe_int(row.get('数量', 0))
                            price = prices.get(style, 0)
                            total_amount += price * qty
                        key = f"{factory_name}-{person}"
                        statistics['by_person'][key] = total_qty
                        statistics['monthly_salary'][key] = round(total_amount, 2)
                
                if '款号' in df.columns and '数量' in df.columns:
                    for style, group in df.groupby('款号'):
                        total_qty = int(group['数量'].sum())
                        price = prices.get(safe_str(style), 0)
                        total_amount = price * total_qty
                        style_key = safe_str(style)
                        if style_key not in statistics['style_summary']:
                            statistics['style_summary'][style_key] = {'quantity': 0, 'amount': 0}
                        statistics['style_summary'][style_key]['quantity'] += total_qty
                        statistics['style_summary'][style_key]['amount'] = round(
                            statistics['style_summary'][style_key]['amount'] + total_amount, 2
                        )
                
                if '颜色' in df.columns and '数量' in df.columns:
                    for color, qty in df.groupby('颜色')['数量'].sum().items():
                        color_key = safe_str(color)
                        statistics['by_color'][color_key] = statistics['by_color'].get(color_key, 0) + int(qty)
                
                if '质量状态' in df.columns:
                    for quality, count in df['质量状态'].value_counts().items():
                        quality_key = safe_str(quality)
                        statistics['by_quality'][quality_key] = statistics['by_quality'].get(quality_key, 0) + int(count)
        
        return jsonify({'success': True, 'statistics': statistics})
    except Exception as e:
        return jsonify({'success': False, 'message': f'统计计算失败: {str(e)}'}), 500

# ---- 下载Excel ----
@app.route('/api/download')
def download_excel():
    if 'username' not in session:
        return jsonify({'success': False, 'message': '请先登录'}), 401
    
    month = request.args.get('month', datetime.now().strftime('%Y-%m'))
    prices = load_prices()
    
    def add_price_amount(records_df):
        df = records_df.copy()
        prices_list = []
        amounts_list = []
        for _, row in df.iterrows():
            style = str(row.get('款号', ''))
            price = prices.get(style, 0)
            qty = row.get('数量', 0)
            try:
                qty = int(qty) if pd.notna(qty) else 0
            except:
                qty = 0
            prices_list.append(price)
            amounts_list.append(round(price * qty, 2))
        df['单价'] = prices_list
        df['金额'] = amounts_list
        return df
    
    if session.get('role') == 'admin':
        all_data = []
        for factory_id, config in FACTORIES.items():
            file_path = config['file']
            if os.path.exists(file_path):
                df = pd.read_excel(file_path)
                if not df.empty and '日期' in df.columns:
                    df = df[df['日期'].astype(str).str.startswith(month)]
                all_data.append(df)
        
        if all_data:
            combined = pd.concat(all_data, ignore_index=True)
            if not combined.empty:
                combined = add_price_amount(combined)
                temp_file = os.path.join(tempfile.gettempdir(), f'all_data_{month}.xlsx')
                combined.to_excel(temp_file, index=False)
                return send_file(temp_file, as_attachment=True, 
                               download_name=f'全部工厂数据_{month}.xlsx')
        
        return jsonify({'success': False, 'message': f'{month}月暂无数据'}), 404
    else:
        factory_id = session.get('factory', 'factory1')
        file_path = get_factory_file(factory_id)
        if file_path and os.path.exists(file_path):
            df = pd.read_excel(file_path)
            current_user = session.get('name', session.get('username'))
            if '人员' in df.columns and not df.empty:
                df = df[df['人员'] == current_user]
            if '日期' in df.columns and not df.empty:
                df = df[df['日期'].astype(str).str.startswith(month)]
            if not df.empty:
                df = add_price_amount(df)
                temp_file = os.path.join(tempfile.gettempdir(), f'my_data_{session.get("username")}_{month}.xlsx')
                df.to_excel(temp_file, index=False)
                return send_file(temp_file, as_attachment=True,
                               download_name=f'{session.get("username")}_生产数据_{month}.xlsx')
            else:
                return jsonify({'success': False, 'message': f'{month}月暂无数据'}), 404
        else:
            return jsonify({'success': False, 'message': '文件不存在'}), 404

# ---- 员工管理（增删改） ----
@app.route('/api/employees', methods=['GET'])
def get_employees():
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        employees = load_employees()
        safe_employees = {}
        for emp_id, emp_info in employees.items():
            factory_id = emp_info.get('factory', '')
            factory_name = FACTORIES[factory_id]['name'] if factory_id in FACTORIES else '全部'
            safe_employees[emp_id] = {
                'name': emp_info.get('name', ''),
                'role': emp_info.get('role', 'employee'),
                'factory': factory_name,
                'factory_id': factory_id
            }
        
        return jsonify({'success': True, 'employees': safe_employees})
    except Exception as e:
        return jsonify({'success': False, 'message': f'获取员工列表失败: {str(e)}'}), 500

@app.route('/api/employees', methods=['POST'])
def add_employee():
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': '未收到数据'}), 400
        
        username = data.get('username', '').strip()
        password = data.get('password', '').strip()
        name = data.get('name', '').strip()
        factory_id = data.get('factory_id', 'factory1')
        
        if not username or not password or not name:
            return jsonify({'success': False, 'message': '用户名、密码、姓名不能为空'}), 400
        
        if len(username) < 3:
            return jsonify({'success': False, 'message': '用户名至少3个字符'}), 400
        
        if len(password) < 4:
            return jsonify({'success': False, 'message': '密码至少4个字符'}), 400
        
        employees = load_employees()
        if username in employees:
            return jsonify({'success': False, 'message': f'用户名 {username} 已存在'}), 400
        
        employees[username] = {
            'password': hashlib.md5(password.encode()).hexdigest(),
            'name': name,
            'role': 'employee',
            'factory': factory_id
        }
        save_employees(employees)
        
        return jsonify({'success': True, 'message': f'员工 {name} 添加成功！'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'添加员工失败: {str(e)}'}), 500

@app.route('/api/employees/<username>', methods=['PUT'])
def update_employee(username):
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        if username == 'admin':
            return jsonify({'success': False, 'message': '不能修改管理员账号'}), 400
        
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': '未收到数据'}), 400
        
        employees = load_employees()
        if username not in employees:
            return jsonify({'success': False, 'message': f'用户 {username} 不存在'}), 404
        
        if 'password' in data and data['password']:
            if len(data['password']) < 4:
                return jsonify({'success': False, 'message': '密码至少4个字符'}), 400
            employees[username]['password'] = hashlib.md5(data['password'].encode()).hexdigest()
        
        if 'name' in data and data['name'].strip():
            employees[username]['name'] = data['name'].strip()
        
        if 'factory_id' in data and data['factory_id']:
            employees[username]['factory'] = data['factory_id']
        
        save_employees(employees)
        
        return jsonify({'success': True, 'message': f'员工 {employees[username]["name"]} 更新成功！'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'更新员工失败: {str(e)}'}), 500

@app.route('/api/employees/<username>', methods=['DELETE'])
def delete_employee(username):
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        if username == 'admin':
            return jsonify({'success': False, 'message': '不能删除管理员账号'}), 400
        
        employees = load_employees()
        if username not in employees:
            return jsonify({'success': False, 'message': f'用户 {username} 不存在'}), 404
        
        del employees[username]
        save_employees(employees)
        
        return jsonify({'success': True, 'message': f'员工 {username} 已删除'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'删除员工失败: {str(e)}'}), 500

# ---- 单价管理API ----
@app.route('/api/prices', methods=['GET'])
def get_prices():
    try:
        if 'username' not in session:
            return jsonify({'success': False, 'message': '请先登录'}), 401
        prices = load_prices()
        return jsonify({'success': True, 'prices': prices})
    except Exception as e:
        return jsonify({'success': False, 'message': f'获取单价失败: {str(e)}'}), 500

@app.route('/api/prices/<style_no>', methods=['GET'])
def get_price(style_no):
    try:
        if 'username' not in session:
            return jsonify({'success': False, 'message': '请先登录'}), 401
        prices = load_prices()
        price = prices.get(style_no, None)
        return jsonify({'success': True, 'style_no': style_no, 'price': price})
    except Exception as e:
        return jsonify({'success': False, 'message': f'获取单价失败: {str(e)}'}), 500

@app.route('/api/prices', methods=['POST'])
def save_price():
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        data = request.get_json()
        if not data or 'style_no' not in data or 'price' not in data:
            return jsonify({'success': False, 'message': '参数不完整'}), 400
        
        style_no = data['style_no'].strip()
        try:
            price = float(data['price'])
        except:
            return jsonify({'success': False, 'message': '单价必须是数字'}), 400
        
        prices = load_prices()
        prices[style_no] = price
        save_prices(prices)
        
        return jsonify({'success': True, 'message': f'款号 {style_no} 单价已设置为 {price} 元'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'保存单价失败: {str(e)}'}), 500

@app.route('/api/prices/<style_no>', methods=['DELETE'])
def delete_price(style_no):
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        prices = load_prices()
        if style_no in prices:
            del prices[style_no]
            save_prices(prices)
            return jsonify({'success': True, 'message': f'款号 {style_no} 单价已删除'})
        else:
            return jsonify({'success': False, 'message': f'款号 {style_no} 不存在'}), 404
    except Exception as e:
        return jsonify({'success': False, 'message': f'删除单价失败: {str(e)}'}), 500

# ---- 工资模板 ----
@app.route('/api/admin/salary-template')
def generate_salary_template():
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        import subprocess
        result = subprocess.run(
            [sys.executable, 'create_salary_template.py'],
            capture_output=True, text=True, encoding='utf-8'
        )
        
        if result.returncode == 0:
            import glob
            template_files = glob.glob('工资计算模板_*.xlsx')
            if template_files:
                latest_template = max(template_files, key=os.path.getctime)
                return jsonify({'success': True, 'message': '工资模板生成成功', 'template_file': latest_template})
            else:
                return jsonify({'success': False, 'message': '未找到生成的模板文件'}), 500
        else:
            return jsonify({'success': False, 'message': '工资模板生成失败', 'error': result.stderr}), 500
    except Exception as e:
        return jsonify({'success': False, 'message': f'生成工资模板失败: {str(e)}'}), 500

@app.route('/api/admin/download-salary-template')
def download_salary_template():
    try:
        if 'username' not in session or session.get('role') != 'admin':
            return jsonify({'success': False, 'message': '权限不足'}), 403
        
        import glob
        template_files = glob.glob('工资计算模板_*.xlsx')
        if not template_files:
            import subprocess
            subprocess.run([sys.executable, 'create_salary_template.py'], capture_output=True)
            template_files = glob.glob('工资计算模板_*.xlsx')
        
        if template_files:
            latest_template = max(template_files, key=os.path.getctime)
            return send_file(latest_template, as_attachment=True)
        else:
            return jsonify({'success': False, 'message': '未找到工资模板文件'}), 404
    except Exception as e:
        return jsonify({'success': False, 'message': f'下载工资模板失败: {str(e)}'}), 500

# ========== 主程序 ==========
if __name__ == '__main__':
    init_excel()
    init_employees()
    init_prices()
    init_backup()
    
    # Railway云部署使用环境变量 PORT，本地默认5000
    port = int(os.environ.get('PORT', 5000))
    
    print("=" * 50)
    print("制造工厂数据录入系统 v3.3")
    print("=" * 50)
    print("系统已初始化完成:")
    print(f"1. 一厂数据文件: {FACTORIES['factory1']['file']}")
    print(f"2. 二厂数据文件: {FACTORIES['factory2']['file']}")
    print(f"3. 员工配置: {EMPLOYEES_FILE}")
    print(f"4. 单价配置: {PRICE_FILE}")
    print(f"5. 备份目录: {BACKUP_DIR}")
    print("\n默认账号:")
    print("管理员: admin / admin123")
    print("一厂员工: f1_emp01 / f1_emp01123 (f1_emp01-f1_emp10)")
    print("二厂员工: f2_emp01 / f2_emp01123 (f2_emp01-f2_emp10)")
    print(f"\n访问地址: http://0.0.0.0:{port}")
    print("按 Ctrl+C 停止服务器")
    print("=" * 50)
    
    try:
        app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)
    except Exception as e:
        print(f"✗ 服务器启动失败: {e}")
        input("按Enter键退出...")
