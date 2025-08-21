import pandas as pd
import math
from collections import defaultdict
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import re
import os
import time
import numpy as np # Added for np.linspace in visualize_packing

def calculate_weight(length, width, count, thickness=32):
    """计算单块格栅重量(kg) - 考虑厚度"""
    volume = length * width * thickness * count / 1e9  # 体积(m³)
    return 7850 * volume  # 钢的密度 7850 kg/m³

def read_excel_data(file_path, sheet_name):
    """从Excel文件读取格栅数据"""
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"成功读取文件: {file_path}")
        print(f"工作表: {sheet_name}")
        print(f"数据行数: {len(df)}")
        
        gratings = []
        
        # 查找关键列
        length_col = None
        width_col = None
        count_col = None
        id_col = None
        el_col = None
        
        # 自动识别列名
        for col in df.columns:
            col_str = str(col).lower()
            if '长度' in col_str or '长' in col_str:
                length_col = col
            elif '宽度' in col_str or '宽' in col_str:
                width_col = col
            elif '数量' in col_str or '块' in col_str:
                count_col = col
            elif '板号' in col_str or '编号' in col_str:
                id_col = col
            elif 'el' in col_str:
                el_col = col
        
        if not all([length_col, width_col, count_col, id_col]):
            print("警告: 无法自动识别必要的列，尝试使用默认列索引")
            # 使用默认列索引
            if len(df.columns) >= 5:
                length_col = df.columns[2]  # 第3列
                width_col = df.columns[3]   # 第4列
                count_col = df.columns[4]   # 第5列
                id_col = df.columns[0]      # 第1列
                el_col = df.columns[7] if len(df.columns) > 7 else None
        
        print(f"识别的列: 长度={length_col}, 宽度={width_col}, 数量={count_col}, 板号={id_col}, EL={el_col}")
        
        # 处理数据行
        for idx, row in df.iterrows():
            try:
                # 跳过标题行和汇总行
                if pd.isna(row[id_col]) or str(row[id_col]).startswith('小计') or str(row[id_col]).startswith('合计'):
                    continue
                
                # 提取数据
                board_id = str(row[id_col]).strip()
                length = int(re.sub(r'[^\d]', '', str(row[length_col])))
                width = int(re.sub(r'[^\d]', '', str(row[width_col])))
                count = int(re.sub(r'[^\d]', '', str(row[count_col])))
                el = str(row[el_col]).strip() if el_col and not pd.isna(row[el_col]) else "未知"
                
                # 验证数据
                if length <= 0 or width <= 0 or count <= 0:
                    continue
                
                # 计算重量（考虑32mm厚度）
                total_weight = calculate_weight(length, width, count, 32)
                single_weight = total_weight / count if count > 0 else 0
                
                # 展开多块格栅
                for _ in range(count):
                    gratings.append({
                        'id': board_id,
                        'length': length,
                        'width': width,
                        'thickness': 32,
                        'weight': single_weight,
                        'el': el,
                        'area': length * width,
                        'volume': length * width * 32
                    })
                    
            except (ValueError, IndexError) as e:
                print(f"解析第{idx+1}行时出错: {row} | {str(e)}")
                continue
        
        return gratings
        
    except Exception as e:
        print(f"读取Excel文件时出错: {str(e)}")
        return []

def fast_pack_gratings(gratings, max_weight=5000):
    """快速高效的二维装箱算法 - 动态调整框架边界，提高利用率"""
    
    print("开始执行装箱算法...")
    start_time = time.time()
    
    # 分析数据，找到最大尺寸
    max_length = max(g['length'] for g in gratings)
    max_width = max(g['width'] for g in gratings)
    
    # 动态确定框架边界，不超过2000*4000
    frame_width = min(max_width, 4000)
    frame_height = min(max_length, 2000)
    
    # 如果最大尺寸超过框架，则使用框架作为边界
    if max_length > frame_height or max_width > frame_width:
        print(f"检测到超大格栅: 最大长度{max_length}mm, 最大宽度{max_width}mm")
        print(f"使用框架边界: {frame_width} × {frame_height} mm")
    
    # 按EL分组
    groups = defaultdict(list)
    for g in gratings:
        groups[g['el']].append(g)
    
    print(f"按EL分组完成，共{len(groups)}个组")
    
    # 打包结果
    packs = []
    pack_id = 1
    
    for el, items in groups.items():
        print(f"正在处理EL组: {el}，共{len(items)}块格栅")
        
        if not items:
            continue
            
        # 按体积降序排序，优先放置大件
        items.sort(key=lambda x: x['volume'], reverse=True)
        
        # 限制每个包裹的最大格栅数量，避免算法卡住
        max_items_per_pack = 50
        
        while items:
            # 初始化新包裹
            pack = {
                'id': f"Pack-{pack_id}",
                'width': 0,
                'height': 0,
                'weight': 0,
                'items': [],
                'positions': [],
                'el': el,
                'utilization': 0,
                'thickness': 32,
                'frame_width': frame_width,
                'frame_height': frame_height
            }
            pack_id += 1
            
            # 使用优化的装箱策略
            current_x = 0
            current_y = 0
            row_height = 0
            items_in_pack = 0
            
            # 尝试添加物品到当前包裹
            i = 0
            while i < len(items) and items_in_pack < max_items_per_pack:
                item = items[i]
                
                # 检查重量约束
                if pack['weight'] + item['weight'] > max_weight:
                    i += 1
                    continue
                
                # 尝试两种旋转方向，优先选择更合适的
                rotations = [
                    (item['length'], item['width']),
                    (item['width'], item['length'])
                ]
                
                # 选择更合适的旋转方向（优先选择能更好利用空间的）
                best_rotation = None
                best_score = float('inf')
                
                for w, h in rotations:
                    if w > frame_width or h > frame_height:
                        continue
                    
                    # 计算放置后的空间利用率
                    new_width = max(pack['width'], current_x + w)
                    new_height = max(pack['height'], current_y + h)
                    score = new_width * new_height
                    
                    if score < best_score:
                        best_score = score
                        best_rotation = (w, h)
                
                if best_rotation is None:
                    i += 1
                    continue
                
                w, h = best_rotation
                
                # 检查是否可以放在当前行
                if current_x + w <= frame_width:
                    # 放在当前行
                    x, y = current_x, current_y
                    current_x += w
                    row_height = max(row_height, h)
                elif current_y + row_height + h <= frame_height:
                    # 换新行
                    current_x = w
                    current_y += row_height
                    row_height = h
                    x, y = 0, current_y
                else:
                    # 无法放置，跳过
                    i += 1
                    continue
                
                # 添加到包裹
                pack['items'].append(item)
                pack['positions'].append((x, y, w, h))
                pack['weight'] += item['weight']
                pack['width'] = max(pack['width'], x + w)
                pack['height'] = max(pack['height'], y + h)
                items_in_pack += 1
                items.pop(i)
            
            # 计算空间利用率
            if pack['items']:
                pack['utilization'] = (pack['width'] * pack['height']) / (frame_width * frame_height)
                packs.append(pack)
                print(f"  完成包裹 {pack['id']}: {len(pack['items'])}块格栅，利用率 {pack['utilization']:.2%}")
    
    end_time = time.time()
    print(f"装箱算法执行完成，耗时: {end_time - start_time:.2f}秒")
    print(f"使用框架边界: {frame_width} × {frame_height} mm")
    
    return packs

def visualize_packing(pack):
    """可视化装箱结果 - 使用动态框架边界"""
    # 使用包裹的框架边界
    frame_width = pack.get('frame_width', 4000)
    frame_height = pack.get('frame_height', 2000)
    
    fig, ax = plt.subplots(figsize=(12, 8))
    ax.set_xlim(0, frame_width)
    ax.set_ylim(0, frame_height)
    ax.set_title(f"包裹: {pack['id']} | EL组: {pack['el']} | 利用率: {pack['utilization']:.2%}")
    ax.set_xlabel("宽度 (mm)")
    ax.set_ylabel("高度 (mm)")
    
    # 绘制框架边界
    frame_rect = patches.Rectangle(
        (0, 0), frame_width, frame_height,
        linewidth=2, edgecolor='gray', facecolor='none', linestyle='-', alpha=0.5
    )
    ax.add_patch(frame_rect)
    
    # 绘制实际包裹边界
    container = patches.Rectangle(
        (0, 0), pack['width'], pack['height'],
        linewidth=2, edgecolor='red', facecolor='none', linestyle='--'
    )
    ax.add_patch(container)
    
    # 绘制每个格栅
    colors = plt.cm.Set3(np.linspace(0, 1, len(pack['items'])))
    for i, (x, y, w, h) in enumerate(pack['positions']):
        rect = patches.Rectangle(
            (x, y), w, h,
            linewidth=1, edgecolor='black', facecolor=colors[i], alpha=0.7
        )
        ax.add_patch(rect)
        
        # 显示板号缩写
        short_id = pack['items'][i]['id'].split('-')[-1] if '-' in pack['items'][i]['id'] else pack['items'][i]['id'][-3:]
        ax.text(x + w/2, y + h/2, short_id, 
                ha='center', va='center', fontsize=8, weight='bold')
    
    # 添加框架信息
    ax.text(frame_width/2, frame_height + 50, 
            f"框架: {frame_width} × {frame_height} mm", 
            ha='center', va='bottom', fontsize=10, weight='bold')
    
    plt.grid(True, alpha=0.3)
    plt.gca().set_aspect('equal', adjustable='box')
    plt.tight_layout()
    plt.show()

def print_packing_summary(packs):
    """打印打包结果摘要"""
    print("\n" + "="*60)
    print("装箱结果摘要")
    print("="*60)
    
    total_weight = 0
    total_items = 0
    total_area = 0
    total_volume = 0
    
    for i, pack in enumerate(packs):
        pack_weight = pack['weight'] / 1000  # 转换为吨
        pack_area = pack['width'] * pack['height'] / 1e6  # 转换为平方米
        pack_volume = pack_area * pack['thickness'] / 1000  # 转换为立方米
        
        frame_width = pack.get('frame_width', 4000)
        frame_height = pack.get('frame_height', 2000)
        
        print(f"\n包裹 #{i + 1}: {pack['id']}")
        print(f"  EL组: {pack['el']}")
        print(f"  框架边界: {frame_width} × {frame_height} mm")
        print(f"  实际尺寸: {pack['width']} × {pack['height']} × {pack['thickness']} mm")
        print(f"  面积: {pack_area:.2f} m²")
        print(f"  体积: {pack_volume:.3f} m³")
        print(f"  重量: {pack_weight:.2f} 吨")
        print(f"  格栅数量: {len(pack['items'])} 块")
        print(f"  空间利用率: {pack['utilization']:.2%}")
        
        total_weight += pack_weight
        total_items += len(pack['items'])
        total_area += pack_area
        total_volume += pack_volume
    
    print("\n" + "="*60)
    print(f"总计: {len(packs)} 个包裹")
    print(f"总重量: {total_weight:.2f} 吨")
    print(f"总格栅: {total_items} 块")
    print(f"总面积: {total_area:.2f} m²")
    print(f"总体积: {total_volume:.3f} m³")
    print("="*60) 

def main():
    """主程序"""
    print("格栅装箱优化算法 V2.0")
    print("="*40)
    
    # 获取用户输入
    while True:
        file_path = input("请输入Excel文件路径 (或按回车使用默认文件): ").strip()
        if not file_path:
            file_path = "生产阿拉木图120E561格栅清单 - 副本.xlsx"
        
        if os.path.exists(file_path):
            break
        else:
            print(f"文件不存在: {file_path}")
    
    while True:
        sheet_name = input("请输入工作表名称 (或按回车使用默认工作表): ").strip()
        if not sheet_name:
            sheet_name = "清单"
        
        try:
            # 验证工作表是否存在
            test_df = pd.read_excel(file_path, sheet_name=sheet_name)
            break
        except:
            print(f"工作表 '{sheet_name}' 不存在，请重新输入")
    
    print(f"\n正在处理文件: {file_path}")
    print(f"工作表: {sheet_name}")
    
    # 读取数据
    gratings = read_excel_data(file_path, sheet_name)
    
    if not gratings:
        print("未找到有效的格栅数据")
        return
    
    print(f"\n成功提取 {len(gratings)} 块格栅")
    
    # 执行快速装箱算法
    print("\n正在执行快速装箱算法...")
    packs = fast_pack_gratings(gratings)
    
    if not packs:
        print("装箱失败")
        return
    
    # 显示结果
    print_packing_summary(packs)
    
    # 自动显示所有装箱可视化图
    print("\n正在生成所有装箱可视化图...")
    for i, pack in enumerate(packs):
        print(f"显示包裹 {i+1}/{len(packs)}: {pack['id']}")
        visualize_packing(pack)
        # 添加短暂延迟，避免图表重叠
        plt.pause(0.5)
    
    print("\n程序执行完成!")

if __name__ == "__main__":
    import numpy as np
    main() 