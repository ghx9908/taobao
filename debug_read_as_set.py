#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试read_as_set函数调用问题
"""

import os

def debug_read_as_set_call():
    """调试read_as_set调用问题"""
    
    print("🔍 调试read_as_set调用问题...")
    
    # 模拟美国taobao.py中的调用
    file_name = "技术风险厄瓜多尔Revista Lideres (Quito, Ecuador).xlsx"
    
    print(f"📁 任务文件名: {file_name}")
    
    # 检查当前目录的results.txt
    current_dir_results = "results.txt"
    print(f"🔍 检查当前目录的results.txt: {current_dir_results}")
    print(f"   存在: {os.path.exists(current_dir_results)}")
    
    # 检查任务目录的results.txt
    task_dir = f"task_data/{file_name.replace('.xlsx', '')}"
    task_results = os.path.join(task_dir, "results.txt")
    print(f"🔍 检查任务目录的results.txt: {task_results}")
    print(f"   存在: {os.path.exists(task_results)}")
    
    # 模拟read_as_set函数的行为
    def simulate_read_as_set(filename="results.txt"):
        """模拟美国taobao.py中的read_as_set函数"""
        print(f"🔧 模拟read_as_set(filename='{filename}')")
        if not os.path.exists(filename):
            print(f"❌ 文件不存在: {filename}")
            return set()
        
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read()
            print(f"✅ 文件存在，内容长度: {len(content)} 字符")
            return set([content])  # 简化处理
    
    # 测试当前目录调用
    print(f"\n🧪 测试1: read_as_set() - 使用默认参数")
    result1 = simulate_read_as_set()
    print(f"   结果: {len(result1)} 条记录")
    
    # 测试任务目录调用
    print(f"\n🧪 测试2: read_as_set('{task_results}') - 指定任务目录文件")
    result2 = simulate_read_as_set(task_results)
    print(f"   结果: {len(result2)} 条记录")
    
    # 分析问题
    print(f"\n📊 问题分析:")
    print(f"   当前目录results.txt存在: {os.path.exists(current_dir_results)}")
    print(f"   任务目录results.txt存在: {os.path.exists(task_results)}")
    
    if not os.path.exists(current_dir_results) and os.path.exists(task_results):
        print(f"❌ 问题确认: 文件在任务目录，但read_as_set()查找当前目录")
        print(f"💡 解决方案: 需要传递正确的文件路径或添加task_filename参数支持")
    
    return result1, result2

def check_file_structure():
    """检查文件结构"""
    print(f"\n📁 文件结构检查:")
    
    # 检查当前目录
    current_files = [f for f in os.listdir('.') if f.endswith('.txt')]
    print(f"   当前目录txt文件: {current_files}")
    
    # 检查任务目录
    task_dirs = [d for d in os.listdir('.') if d.startswith('task_data')]
    print(f"   任务目录: {task_dirs}")
    
    for task_dir in task_dirs:
        if os.path.isdir(task_dir):
            task_files = os.listdir(task_dir)
            print(f"   {task_dir}/: {task_files}")

if __name__ == "__main__":
    debug_read_as_set_call()
    check_file_structure()
