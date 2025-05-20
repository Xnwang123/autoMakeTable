import pandas as pd
import re
from collections import defaultdict

def runUtilics(upload_file,to_one = 2,to_two = 3,nx_one = 2,nx_two=3,night = 2):
    # 时间段人数配置（新增配置字典）
    TIME_SLOT_REQUIREMENT = {
        "1-2节": to_one,    # 早上第一班
        "3-5节": to_two,    # 早上第二班
        "6-7节": nx_one,    # 下午第一班
        "8-9节": nx_two,    # 下午第二班
        "10-11节": night  # 晚上（至少2人，可灵活）
    }

    # 白天时间段定义（用于优先级判断）
    DAYTIME_SLOTS = ["1-2节", "3-5节", "6-7节", "8-9节"]

    # 读取Excel数据
    #data = pd.read_excel("py_auto_table/singal.xlsx")
    data = pd.read_excel(upload_file)
    df = pd.DataFrame(data)

    # 数据预处理
    schedule = {}
    all_people = set()
    for idx, row in df.iterrows():
        time_slot = row["时间"]
        schedule[time_slot] = {}
        for day in ["星期一", "星期二", "星期三", "星期四", "星期五"]:
            people = list(set(row[day].split(",")))  # 去重处理
            schedule[time_slot][day] = people
            all_people.update(people)

    # 初始化统计信息
    person_stats = defaultdict(lambda: {
        "total": 0,  # 周总次数
        "daily": defaultdict(int),  # 每日次数
        "isAssignment":False   #判断当前人员是否已经分配
    })
    assignments = {day: {slot: [] for slot in schedule} for day in ["星期一", "星期二", "星期三", "星期四", "星期五"]}
    warning_log = []  # 警告信息存储

    def can_assign(person, day):
        """检查分配条件"""
        return (person_stats[person]["total"] < 2 and 
                person_stats[person]["daily"][day] < 2)

    # 分配函数（新增时间段人数参数）
    def assign_people(day, time_slot, required):
        candidates = schedule[time_slot][day]
        if not candidates:
            warning_log.append(f"{day}-{time_slot}: 无可用值班人员")
            return 0
        
        # 排序策略：优先未分配人员，其次次数少的
        candidates_sorted = sorted(
            candidates,
            key=lambda x: (
                person_stats[x]["total"],
                len([d for d in person_stats[x]["daily"] if person_stats[x]["daily"][d] > 0])
            )
        )
        
        assigned = []
        for _ in range(required):
            found = False
            for p in candidates_sorted:
                if p in assigned: continue  # 禁止重复
                if can_assign(p, day):
                    assignments[day][time_slot].append(p)
                    person_stats[p]["total"] += 1
                    person_stats[p]["daily"][day] += 1
                    person_stats[p]["isAssignment"] = True
                    assigned.append(p)
                    found = True
                    break
            if not found:
                break  # 无法找到更多人员
        
        # 记录警告信息
        if len(assigned) < required:
            missing = required - len(assigned)
            warning_log.append(
                f"{day}-{time_slot}: 需要{required}人，实际分配{len(assigned)}人 "
                f"(候选：{candidates_sorted})"
            )
        return len(assigned)
    
    # 第一阶段：优先分配白天时段
    unassigned = set(all_people)
    for day in ["星期一", "星期二", "星期三", "星期四", "星期五"]:
        # 先处理白天时段
        for time_slot in DAYTIME_SLOTS:
            required = TIME_SLOT_REQUIREMENT[time_slot]
            assign_people(day, time_slot, required)
    
        # 最后处理晚上时段（10-11节）
        time_slot = "10-11节"
        required = TIME_SLOT_REQUIREMENT[time_slot]
        assigned_count = assign_people(day, time_slot, required)
    
        # 晚上时段灵活处理：如果分配不足但有人值班，不强制补足
        if 0 < assigned_count < required:
            warning_log.append(f"{day}-{time_slot}: 晚间时段分配不足（{assigned_count}/{required}）")

    # 第二阶段：强制分配剩余人员
    while unassigned:
        person = next(iter(unassigned))
        assigned_flag = False
    
        # 优先尝试白天时段
        for day in ["星期一", "星期二", "星期三", "星期四", "星期五"]:
            for time_slot in DAYTIME_SLOTS:
                if person not in schedule[time_slot][day]:
                    continue
                if (can_assign(person, day) and 
                    len(assignments[day][time_slot]) < TIME_SLOT_REQUIREMENT[time_slot] and
                    person not in assignments[day][time_slot]):
                    assignments[day][time_slot].append(person)
                    person_stats[person]["total"] += 1
                    person_stats[person]["daily"][day] += 1
                    assigned_flag = True
                    unassigned.remove(person)
                    break
            if assigned_flag:
                break
        
        # 如果白天无法分配，尝试晚上时段
        if not assigned_flag:
            for day in ["星期一", "星期二", "星期三", "星期四", "星期五"]:
                time_slot = "10-11节"
                if person not in schedule[time_slot][day]:
                    continue
                if (can_assign(person, day) and 
                    len(assignments[day][time_slot]) < TIME_SLOT_REQUIREMENT[time_slot] and
                    person not in assignments[day][time_slot]):
                    assignments[day][time_slot].append(person)
                    person_stats[person]["total"] += 1
                    person_stats[person]["daily"][day] += 1
                    assigned_flag = True
                    unassigned.remove(person)
                    break
            if assigned_flag:
                break
        #未分配人员发出警告！
        if not assigned_flag:
            if person_stats[person]["isAssignment"] == False:
                warning_log.append(f"无法为 {person} 安排值班，请检查可用时间")
            unassigned.remove(person)
    
    # 创建Excel写入对象
    #writer = pd.ExcelWriter('py_auto_table/排班结果.xlsx', engine='openpyxl')
    
    # 格式化输出
    result = []
    for day in ["星期一", "星期二", "星期三", "星期四", "星期五"]:
        day_data = {"日期": day}
        for time_slot in schedule:
            assigned = assignments[day][time_slot]
            required = TIME_SLOT_REQUIREMENT[time_slot]
            status = ",".join(assigned)
            # 添加状态标记
            if len(assigned) < required:
                if time_slot == "10-11节" and len(assigned) >= 1:
                    status += "（晚间人手不足）"
                else:
                    status += f"（需补{required-len(assigned)}人）"
            day_data[time_slot] = status
        result.append(day_data)

    result_df = pd.DataFrame(result).set_index("日期")
    # 输出主排班表
    #result_df.T.to_excel(writer, sheet_name='排班表')
    #创建人员统计表
    stats_data = []
    for person in sorted(all_people):
        stats = person_stats[person]
        days = ", ".join([d for d in stats["daily"] if stats["daily"][d] > 0])
        stats_data.append({
            "姓名": person,
            "目前安排值班次数": stats["total"],
            "当周需要值班次数": len(stats["daily"]),
            "具体日期": days,
            "是否分配":stats["isAssignment"]
        })
    stats_df = pd.DataFrame(stats_data)

    # 创建警告日志表
    if warning_log:
        warn_df = pd.DataFrame({"警告信息": warning_log})
        

    # 保存文件
    #writer.close()
    return result_df,stats_df,warn_df

#创建企业对接表函数
def runMakeEnterprise(df_initTable,df_worker):
    def makeFile(file_upload):
        # 读取指定工作表
        df = pd.read_excel(file_upload)

        # 选择需要的列
        selected_columns = [
            "场次", "单位名称", "单位性质", "联系人", "联系电话", 
            "招聘地点", "星期", "宣讲会开始时间", "宣讲会结束时间"
        ]
        result_df = df[selected_columns]

        # 合并"宣讲会开始时间"和"宣讲会结束时间"为"宣讲时间"
        result_df["宣讲时间"] = (
            result_df["宣讲会开始时间"].astype(str).str.split().str[0]  # 提取日期部分（YYYY-MM-DD）
            + " " 
            + result_df["宣讲会开始时间"].astype(str).str.split().str[1].str[:5]  # 提取开始时间（HH:MM）
            + "-" 
            + result_df["宣讲会结束时间"].astype(str).str.split().str[1].str[:5]  # 提取结束时间（HH:MM）
        )
        # 删除原来的两列
        result_df = result_df.drop(columns=["宣讲会开始时间", "宣讲会结束时间"])

        # 调整列顺序（可选）
        result_df = result_df[[
            "场次", "单位名称", "单位性质", "联系人", "联系电话", 
            "招聘地点", "星期", "宣讲时间"
        ]]
        
        return result_df


    # 读取课程表CSV文件并清洗数据
    def load_schedule_data(csv_path):
        # 处理BOM字符并读取CSV
        schedule_df = pd.read_csv(csv_path, encoding='utf-8-sig', index_col=0)
        
        # 清洗助理姓名中的备注（如括号内容）
        def clean_name(cell):
            if pd.isna(cell):
                return []
            # 去除括号及内容，并分割姓名
            cleaned = re.sub(r'（.*?）', '', str(cell))
            return [name.strip() for name in cleaned.split(',') if name.strip()]
        
        # 对每个单元格应用清洗
        for col in schedule_df.columns:
            schedule_df[col] = schedule_df[col].apply(clean_name)
        
        return schedule_df

    # 映射时间段到课程节数
    time_slot_mapping = {
        '09:00-10:00': '1-2节',
        '10:30-12:00': '3-5节',
        '14:30-15:30': '6-7节',
        '16:00-17:00': '8-9节',
        '18:30-23:59': '10-11节'  # 假设18:30后统一映射
    }

    # 从宣讲时间解析时间段
    def parse_time_slot(time_str):
        try:
            _, time_range = time_str.split(' ')
            start_time, end_time = time_range.split('-')
            # 标准化时间格式
            start_time = f"{start_time[:2]}:{start_time[3:5]}"
            end_time = f"{end_time[:2]}:{end_time[3:5]}"
            
            # 匹配预定义时间段
            for slot, section in time_slot_mapping.items():
                slot_start, slot_end = slot.split('-')
                if slot_start <= start_time < slot_end:
                    return section
            # 处理18:30之后的情况
            if start_time >= '18:30':
                return '10-11节'
            return None
        except:
            return None
        
    # 加载课程表数据
    schedule_df = load_schedule_data(df_worker)
        
    # 从生成的初版企业对接表中读取已有的宣讲会数据
    extracted_df = makeFile(df_initTable)
        
    # 为每场宣讲会分配助理
    def assign_assistants(row):
        # 如果招聘地点包含"麦庐园"，则跳过分配
        if '麦庐园' in str(row['招聘地点']):
            return " "
        day = row['星期']
        time_slot = parse_time_slot(row['宣讲时间'])
            
        if not time_slot or day not in schedule_df.columns:
            return "需手动分配"  # 异常情况标记
            
        assistants = schedule_df.loc[time_slot, day]
        # 至少保证两位助理
        if len(assistants) >= 2:
            return ', '.join(assistants[:2])
        else:
            return ', '.join(assistants + ['备用助理']*(2-len(assistants)))
        
    # 添加对接助理列
    extracted_df['对接助理'] = extracted_df.apply(assign_assistants, axis=1)
    return extracted_df