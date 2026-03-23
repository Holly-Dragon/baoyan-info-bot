import requests
import json
import time
import os
import pandas as pd
from datetime import datetime

def get_and_update_college_info(excel_file="2026院校信息.xlsx"):
    """
    从保研网API获取院校信息，计算倒计时并更新到Excel中
    """
    url = "http://api.baoyanwang.com.cn/api/v1/articles"
    # 模拟浏览器，避免被接口拦截
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36 Edg/129.0.0.0"
    }

    # 加载已有的Excel数据
    old_data = []
    if os.path.exists(excel_file):
        try:
            df_old = pd.read_excel(excel_file)
            df_old = df_old.fillna('')
            old_data = df_old.to_dict('records')
            print(f"成功加载已有本地文件，共 {len(old_data)} 条记录。")
        except Exception as e:
            print(f"读取旧Excel出错，将重新创建：{e}")

    # 获取当前日期做对比
    now = datetime.now()
    all_data_map = {item.get('通知链接'): item for item in old_data}
    
    print("\n=== 开始拉取 API 更新数据 ===")
    page = 1
    max_pages_without_new = 3 # 当连续3页没有新数据时停止抓取
    pages_without_new_count = 0
    has_any_updates_overall = False
    
    while True:
        params = {
            "page": page,
            "size": 25,
            "category": "保研信息",
            "all": 1
        }
        
        try:
            response = requests.get(url, params=params, headers=headers)
            response.raise_for_status()
            content = json.loads(response.text).get('result', {}).get('content', [])
            
            if not content:
                print("未获取到更多数据，结束抓取。")
                break
                
            has_new_items_in_page = False
            for item in content:
                # 过滤条件：2026年的项目
                year_info = item.get('year', 0)
                origin_title = item.get('title', '')
                sign_up_end_str = item.get('sign_up_end', '')
                
                # 若年份标记为2026，或标题/截止时间中包含2026，均视为2026年项目
                if year_info == 2026 or '2026' in origin_title or '2026' in sign_up_end_str:
                    office_url = item.get('office_url', '无链接')
                    
                    # 信息清洗和提取
                    school = item.get('college', '')
                    academy = item.get('academy', '')
                    major = item.get('major', '')
                    project = origin_title
                    
                    # 如果学校字段为空，尝试从标题【】中提取
                    if '【' in project and '】' in project:
                        parts = project.split('】')
                        if not school:
                            school = parts[0].replace('【', '').strip()
                        if '——' in parts[1]:
                            academy_guess = parts[1].split('——')[1].strip()
                            if not academy:
                                academy = academy_guess.strip()
                    
                    if not school:
                        school = '未知院校'
                    
                    if not academy:
                        academy = '未知学院'
                    
                    # 检查是否是新数据或关键字段发生了更新
                    is_new_or_updated = False
                    if office_url not in all_data_map:
                        is_new_or_updated = True
                    else:
                        old_item = all_data_map[office_url]
                        if old_item.get('招生项目') != project or old_item.get('申请截止时间') != sign_up_end_str:
                            is_new_or_updated = True
                            
                    if is_new_or_updated:
                        has_new_items_in_page = True
                        has_any_updates_overall = True
                        
                        row = {
                            '更新时间': now.strftime("%Y-%m-%d %H:%M:%S"),
                            '院校发布时间': item.get('updated_time', '未知'),
                            '学校': school,
                            '学院': academy,
                            '专业': major,
                            '招生项目': project,
                            '通知链接': office_url,
                            '申请截止时间': sign_up_end_str
                        }
                        
                        # 直接覆盖或新增数据，未更新的数据保留原有的“更新时间”不变
                        if office_url != '无链接':
                            all_data_map[office_url] = row

            print(f"已处理第 {page} 页。")
            if not has_new_items_in_page:
                pages_without_new_count += 1
            else:
                pages_without_new_count = 0
                
            # 当连续若干页没有新增或更新时，停止获取。节约资源并符合抓取规律（最新更新会排在前面）。
            if pages_without_new_count >= max_pages_without_new:
                print("连续多页无新数据，终止后续页面抓取。")
                break
                
            page += 1
            time.sleep(1) # 休眠防止频繁请求被封
            
        except requests.exceptions.RequestException as e:
            print(f"请求异常，抓取终止：{e}")
            break
        except Exception as e:
            print(f"解析出错：{e}")
            break

    final_list = list(all_data_map.values())
    
    # 如果没有任何更新，且本地已有文件，则不重写 Excel 以节约性能
    if not has_any_updates_overall and os.path.exists(excel_file):
        print(f"\n当前汇总 {len(final_list)} 项 2026 院校信息。")
        print("未发现新数据或需要更新的数据，跳过主 Excel 重写。")
        return final_list

    print(f"\n当前汇总 {len(final_list)} 项 2026 院校信息，准备写入/更新 Excel...")
    
    # 转为 DataFrame 保存
    if final_list:
        df_to_save = pd.DataFrame(final_list)
        # 按照需要的字段排序
        columns = ['更新时间', '院校发布时间', '学校', '学院', '专业', '招生项目', '通知链接', '申请截止时间', '申请倒计时', '是否截止']
        # 补齐如果有缺漏字段避免报错
        for col in columns:
            if col not in df_to_save.columns:
                df_to_save[col] = ''
        df_to_save = df_to_save[columns]
        
        # 按院校发布时间由早到晚排序，并重置索引以便接下来注入 Excel 公式
        df_to_save = df_to_save.sort_values(by='院校发布时间', ascending=True).reset_index(drop=True)
        
        # 使用 Excel 内部函数计算申请倒计时和是否截止（利用 H 列：申请截止时间）
        # 保证打开 Excel 文档时根据系统环境动态更新
        # 采用向量化赋值取代 for 循环，大幅提升万级数据量的处理速度
        row_indices = df_to_save.index + 2  # 行号（包含表头，从2开始）
        df_to_save['申请倒计时'] = [f'=IFERROR(IF(VALUE(LEFT(H{idx},10))<TODAY(), "已截止", VALUE(LEFT(H{idx},10))-TODAY() & "天"), "未知")' for idx in row_indices]
        df_to_save['是否截止'] = [f'=IFERROR(IF(VALUE(LEFT(H{idx},10))<TODAY(), "是", "否"), "未知")' for idx in row_indices]
        
        df_to_save.to_excel(excel_file, index=False)
        print(f"\n成功写入 {len(final_list)} 条数据至：{excel_file}")
    else:
        print("\n未筛选到符合条件的数据。")

    return final_list

def export_categorized_excel(data_list, output_file="2026院校信息_分类.xlsx"):
    """
    根据专业将数据划分到不同的 sheet 里面
    """
    if not data_list:
        return
        
    print(f"\n=== 开始生成分类表格：{output_file} ===")
    
    categories = {
        "理工农医": [],
        "经管法学": [],
        "人文社科与艺术": [],
        "单校": []
    }
    
    # 关键词定义
    kw_jingguan = ['经济', '金融', '管理', '商', '法学', '法律', '财', '审计', '会计', '税务', '保险', '行政', '政治', '公共管理', '统计']
    kw_renwen = ['文学', '中文', '外语', '历史', '史学', '哲学', '艺术', '语言', '翻译', '新闻', '传媒', '传播', '社会', '设计', '音乐', '美术', '戏剧', '影视', '体育', '心理', '教育', '马克思', '国际']
    kw_ligong = ['工学', '农学', '医学', '计算', '软件', '电子', '信息', '通信', '网络', '系统', '物理', '化学', '数学', '生物', '材料', '机械', '土木', '航空', '航天', '动力', '电气', '自动', '环境', '地理', '地质', '海洋', '药', '护理', '卫生', '光', '数据', '智能', '制造', '测控', '安全', '技术', '科学', '工程']
    
    for row in data_list:
        major = row.get('专业', '')
        
        # 1. 单校情况：专业为空，或填写了“不限”、“无要求”之类
        if not major or major in ['不限', '无限制', '全校']:
            categories["单校"].append(row)
            continue
            
        matched_categories = set()
        
        # 2. 依次匹配分类，支持一个项目匹配多分类
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
        
    # 写入多个 sheet 到新的 Excel 文件中
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        columns = ['更新时间', '院校发布时间', '学校', '学院', '专业', '招生项目', '通知链接', '申请截止时间', '申请倒计时', '是否截止']
        
        for sheet_name, rows in categories.items():
            if rows:
                df = pd.DataFrame(rows)
                for col in columns:
                    if col not in df.columns:
                        df[col] = ''
                df = df[columns]
                
                df = df.sort_values(by='院校发布时间', ascending=True).reset_index(drop=True)
                
                # 同理注入更新倒计时的函数（利用 H 列）
                row_indices = df.index + 2
                df['申请倒计时'] = [f'=IFERROR(IF(VALUE(LEFT(H{idx},10))<TODAY(), "已截止", VALUE(LEFT(H{idx},10))-TODAY() & "天"), "未知")' for idx in row_indices]
                df['是否截止'] = [f'=IFERROR(IF(VALUE(LEFT(H{idx},10))<TODAY(), "是", "否"), "未知")' for idx in row_indices]
                
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # 即使分类目前为空，也许以后会有数据，这里先创建一个带表头的空 sheet
                pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet_name, index=False)
                
    print(f"成功导出分类表格，各类别统计如下：")
    print(f"  - 理工农医       : {len(categories['理工农医'])} 条")
    print(f"  - 经管法学       : {len(categories['经管法学'])} 条")
    print(f"  - 人文社科与艺术 : {len(categories['人文社科与艺术'])} 条")
    print(f"  - 单校         : {len(categories['单校'])} 条")

if __name__ == "__main__":
    result_data = get_and_update_college_info()
    if result_data:
        export_categorized_excel(result_data)