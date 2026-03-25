import requests
import json
import time
import os
import pandas as pd
from datetime import datetime

def format_date_str(date_str):
    """尝试将日期字符串格式化为 'x年x月x日'"""
    if pd.isna(date_str) or not str(date_str).strip() or str(date_str) == '未知':
        return '未知'
    try:
        dt = pd.to_datetime(str(date_str)[:10])
        return f"{dt.year}年{dt.month}月{dt.day}日"
    except Exception:
        return str(date_str)

def get_and_update_college_info(excel_file="2026院校信息.xlsx", update_file="2026院校信息_更新.xlsx"):
    """
    从保研网API获取院校信息，并只将增量（新/变动）数据输出到更新文件
    """
    url = "http://api.baoyanwang.com.cn/api/v1/articles"
    # 模拟浏览器，避免被接口拦截
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36 Edg/129.0.0.0"
    }

    # 加载已有的Excel数据作为比对基准
    old_data = []
    if os.path.exists(excel_file):
        try:
            df_old = pd.read_excel(excel_file)
            df_old = df_old.fillna('')
            old_data = df_old.to_dict('records')
            print(f"成功加载已有本地基准文件，共 {len(old_data)} 条记录用于对比。")
        except Exception as e:
            print(f"读取旧Excel基准出错：{e}")

    now = datetime.now()
    all_data_map = {item.get('通知链接'): item for item in old_data}
    new_data_list = []
    
    print("\n=== 开始拉取 API 更新数据 ===")
    page = 1
    max_pages_without_new = 3
    pages_without_new_count = 0
    
    while True:
        params = {"page": page, "size": 25, "category": "保研信息", "all": 1}
        try:
            response = requests.get(url, params=params, headers=headers)
            response.raise_for_status()
            content = json.loads(response.text).get('result', {}).get('content', [])
            
            if not content:
                print("未获取到更多数据，结束抓取。")
                break
                
            has_new_items_in_page = False
            for item in content:
                year_info = item.get('year', 0)
                origin_title = item.get('title', '')
                sign_up_end_str = item.get('sign_up_end', '')
                
                if year_info == 2026 or '2026' in origin_title or '2026' in sign_up_end_str:
                    office_url = item.get('office_url', '无链接')
                    
                    school = item.get('college', '')
                    academy = item.get('academy', '')
                    major = item.get('major', '')
                    project = origin_title
                    
                    if '【' in project and '】' in project:
                        parts = project.split('】')
                        if not school:
                            school = parts[0].replace('【', '').strip()
                        if '——' in parts[1] if len(parts)>1 else False:
                            academy_guess = list(filter(None, parts[1].split('——')))[-1].strip()
                            if not academy:
                                academy = academy_guess.strip()
                                
                    if not school: school = '未知院校'
                    if not academy: academy = '未知学院'
                    
                    formatted_end_str = format_date_str(sign_up_end_str)
                    
                    is_new_or_updated = False
                    if office_url not in all_data_map:
                        is_new_or_updated = True
                    else:
                        old_item = all_data_map[office_url]
                        # 对旧数据的日期也进行一样的格式化后再做对比，避免因单改时间格式导致所有历史数据被误判为“更新”
                        old_end_str = format_date_str(old_item.get('申请截止时间'))
                        if old_item.get('招生项目') != project or old_end_str != formatted_end_str:
                            is_new_or_updated = True
                            
                    if is_new_or_updated:
                        has_new_items_in_page = True
                        row = {
                            '更新时间': now.strftime("%Y-%m-%d %H:%M:%S"),
                            '院校报名信息发布时间': format_date_str(item.get('updated_time', '未知')),
                            '学校': school,
                            '学院': academy,
                            '专业': major,
                            '招生项目': project,
                            '通知链接': office_url,
                            '申请截止时间': formatted_end_str
                        }
                        if office_url != '无链接':
                            new_data_list.append(row)
                            all_data_map[office_url] = row

            print(f"已处理第 {page} 页。")
            if not has_new_items_in_page:
                pages_without_new_count += 1
            else:
                pages_without_new_count = 0
                
            if pages_without_new_count >= max_pages_without_new:
                print("连续多页无新数据，终止后续页面抓取。")
                break
                
            page += 1
            time.sleep(1)
            
        except Exception as e:
            print(f"抓取异常：{e}")
            break

    if new_data_list:
        print(f"\n发现 {len(new_data_list)} 条最新更新数据，准备写入：{update_file}")
        export_categorized_excel(new_data_list, update_file)
    else:
        print("\n未发现需要更新的数据。")
        pass

def export_categorized_excel(data_list, output_file="2026院校信息_分类.xlsx"):
    if not data_list:
        return
        
    print(f"\n=== 开始生成分类表格：{output_file} ===")
    categories = {
        "理工农医": [],
        "经管法学": [],
        "人文社科与艺术": [],
        "单校": []
    }
    
    kw_jingguan = ['经济', '金融', '管理', '商', '法学', '法律', '财', '审计', '会计', '税务', '保险', '行政', '政治', '公共管理', '统计']
    kw_renwen = ['文学', '中文', '外语', '历史', '史学', '哲学', '艺术', '语言', '翻译', '新闻', '传媒', '传播', '社会', '设计', '音乐', '美术', '戏剧', '影视', '体育', '心理', '教育', '马克思', '国际']
    kw_ligong = ['工学', '农学', '医学', '计算', '软件', '电子', '信息', '通信', '网络', '系统', '物理', '化学', '数学', '生物', '材料', '机械', '土木', '航空', '航天', '动力', '电气', '自动', '环境', '地理', '地质', '海洋', '药', '护理', '卫生', '光', '数据', '智能', '制造', '测控', '安全', '技术', '科学', '工程']
    
    for row in data_list:
        major = row.get('专业', '')
        if not major or major in ['不限', '无限制', '全校']:
            categories["单校"].append(row)
            continue
            
        matched_categories = set()
        for kw in kw_jingguan:
            if kw in major:
                matched_categories.add("经管法学")
                break
        for kw in kw_renwen:
            if kw in major:
                matched_categories.add("人文社科与艺术")
                break
        for kw in kw_ligong:
            if kw in major:
                matched_categories.add("理工农医")
                break
        
        if not matched_categories:
            matched_categories.add("单校")
            
        for category in matched_categories:
            categories[category].append(row)
        
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        columns = ['更新时间', '院校报名信息发布时间', '学校', '学院', '专业', '招生项目', '通知链接', '申请截止时间', '申请倒计时', '是否截止']
        for sheet_name, rows in categories.items():
            if rows:
                df = pd.DataFrame(rows)
                for col in columns:
                    if col not in df.columns:
                        df[col] = ''
                df = df[columns]
                
                # 按照 '院校报名信息发布时间' 排序
                df['temp_date'] = pd.to_datetime(df['院校报名信息发布时间'].str.extract(r'(\d+年\d+月\d+日)')[0], format='%Y年%m月%d日', errors='coerce')
                df = df.sort_values(by='temp_date', ascending=True, na_position='last').drop(columns=['temp_date']).reset_index(drop=True)
                
                row_indices = df.index + 2
                df['申请倒计时'] = [f'=IFERROR(IF(DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{idx},"年","-"),"月","-"),"日",""))<TODAY(), "已截止", DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{idx},"年","-"),"月","-"),"日",""))-TODAY() & "天"), "未知")' for idx in row_indices]
                df['是否截止'] = [f'=IFERROR(IF(DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{idx},"年","-"),"月","-"),"日",""))<TODAY(), "是", "否"), "未知")' for idx in row_indices]
                
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet_name, index=False)
                
    print(f"成功导出分类表格，各类别统计如下：")
    print(f"  - 理工农医       : {len(categories['理工农医'])} 条")
    print(f"  - 经管法学       : {len(categories['经管法学'])} 条")
    print(f"  - 人文社科与艺术 : {len(categories['人文社科与艺术'])} 条")
    print(f"  - 单校         : {len(categories['单校'])} 条")

if __name__ == "__main__":
    get_and_update_college_info()
