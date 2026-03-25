import os
import pandas as pd
from info_get import export_categorized_excel, format_date_str

def merge_and_output(main_file="2026院校信息.xlsx", update_file="2026院校信息_更新.xlsx", output_categorized="2026院校信息_分类.xlsx"):
    """
    读取更新文件，将其合并写入到主文件中，并重新生成分类版 Excel。
    """
    if not os.path.exists(update_file):
        print(f"未找到更新文件 `{update_file}`，可能没有新数据。合并流程结束。")
        return
        
    print(f"\n=== 开始读取并合并更新数据 ===")
    
    all_data_map = {}
    
    # 读入主文件（如果存在）
    if os.path.exists(main_file):
        df_main = pd.read_excel(main_file)
        df_main = df_main.fillna('')
        
        # # 将已有的旧字段名替换成新字段名（处理旧表格包含的问题）
        # if '院校发布时间' in df_main.columns:
        #     df_main.rename(columns={'院校发布时间': '院校报名信息发布时间'}, inplace=True)
            
        main_data = df_main.to_dict('records')
        for idx, item in enumerate(main_data):
            # 兼容旧数据的日期格式化
            item['院校报名信息发布时间'] = format_date_str(item.get('院校报名信息发布时间', ''))
            item['申请截止时间'] = format_date_str(item.get('申请截止时间', ''))
            link = item.get('通知链接')
            if link and link != '无链接':
                all_data_map[link] = item
            else:
                # 确保无链接的历史数据不被覆盖丢失
                all_data_map[f"old_no_url_{idx}"] = item
            
    # 读入更新文件所有 sheets 并合并
    # read_excel 返回 dict of DataFrames 如果 sheet_name=None
    dfs_update = pd.read_excel(update_file, sheet_name=None)
    update_count = 0
    
    # 建立一个去重集合，避免同一个对象如果在_更新中有好几个分类被重复加入（如果是无链接项）
    seen_updates = set()
    
    for sheet_name, df_update in dfs_update.items():
        if df_update.empty:
            continue
        df_update = df_update.fillna('')
        update_data = df_update.to_dict('records')
        
        for idx, item in enumerate(update_data):
            link = item.get('通知链接')
            proj = item.get('招生项目')
            school = item.get('学校')
            uniq_key = f"{link}_{proj}_{school}"
            
            if uniq_key in seen_updates:
                continue
            seen_updates.add(uniq_key)
            update_count += 1
            
            if link and link != '无链接':
                all_data_map[link] = item
            else:
                all_data_map[f"new_no_url_{sheet_name}_{idx}"] = item
            
    print(f"读取到约 {update_count} 条更新（含交叉项）待合并。")
        
    final_list = list(all_data_map.values())
    print(f"合并后当前汇总数据共 {len(final_list)} 项，准备写入 Excel...")
    
    df_merged = pd.DataFrame(final_list)
    
    # 构建最终需要的字段及公式注入
    columns = ['更新时间', '院校报名信息发布时间', '学校', '学院', '专业', '招生项目', '通知链接', '申请截止时间', '申请倒计时', '是否截止']
    for col in columns:
        if col not in df_merged.columns:
            df_merged[col] = ''
    df_merged = df_merged[columns]
    
    # 确保根据格式化后的 'x年x月x日' 日期从早到晚排序
    df_merged['temp_date'] = pd.to_datetime(df_merged['院校报名信息发布时间'].str.extract(r'(\d+年\d+月\d+日)')[0], format='%Y年%m月%d日', errors='coerce')
    df_merged = df_merged.sort_values(by='temp_date', ascending=True, na_position='last').drop(columns=['temp_date']).reset_index(drop=True)
    
    # 注入截止到期日和状态的 Excel 函数公式（以H列 "x年x月x日" 作为解析源）
    row_indices = df_merged.index + 2
    df_merged['申请倒计时'] = [f'=IFERROR(IF(DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{idx},"年","-"),"月","-"),"日",""))<TODAY(), "已截止", DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{idx},"年","-"),"月","-"),"日",""))-TODAY() & "天"), "未知")' for idx in row_indices]
    df_merged['是否截止'] = [f'=IFERROR(IF(DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{idx},"年","-"),"月","-"),"日",""))<TODAY(), "是", "否"), "未知")' for idx in row_indices]
    
    df_merged.to_excel(main_file, index=False)
    print(f"成功更新主文件：{main_file}")
    
    # 4. 生成分类文件
    export_categorized_excel(final_list, output_categorized)

if __name__ == "__main__":
    merge_and_output()
