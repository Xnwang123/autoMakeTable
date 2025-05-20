import pandas as pd
from collections import defaultdict

# 时间段人数配置（新增配置字典）
TIME_SLOT_REQUIREMENT = {
    "1-2节": 2,    # 早上第一班
    "3-5节": 3,    # 早上第二班
    "6-7节": 2,    # 下午第一班
    "8-9节": 3,    # 下午第二班
    "10-11节": 2  # 晚上（至少2人，可灵活）
}

# 白天时间段定义（用于优先级判断）
DAYTIME_SLOTS = ["1-2节", "3-5节", "6-7节", "8-9节"]

# 读取Excel数据
data = pd.read_excel("py_auto_table/singal.xlsx")
df = pd.DataFrame(data)

# 数据预处理
schedule = {}
all_people = set()
person_availability = defaultdict(int)  # 新增：统计人员出现次数

for idx, row in df.iterrows():
    time_slot = row["时间"]
    schedule[time_slot] = {}
    for day in ["周一", "周二", "周三", "周四", "周五"]:
        people = list(set(row[day].split(",")))
        schedule[time_slot][day] = people
        all_people.update(people)
        # 统计人员出现次数
        for p in people:
            person_availability[p] += 1

# 初始化统计信息
person_stats = defaultdict(lambda: {
    "total": 0,  # 周总次数
    "daily": defaultdict(int)  # 每日次数
})
assignments = {day: {slot: [] for slot in schedule} for day in ["周一", "周二", "周三", "周四", "周五"]}
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
    
    # 修改排序策略：优先可用次数少的人员
    candidates_sorted = sorted(
        candidates,
        key=lambda x: (
            person_availability[x],  # 可用次数越少优先级越高
            person_stats[x]["total"],
            len([d for d in person_stats[x]["daily"] if person_stats[x]["daily"][d] > 0])
        )
    )
    
    assigned = []
    for _ in range(required):
        found = False
        for p in candidates_sorted:
            if p in assigned: continue
            if can_assign(p, day):
                assignments[day][time_slot].append(p)
                person_stats[p]["total"] += 1
                person_stats[p]["daily"][day] += 1
                assigned.append(p)
                found = True
                break
        if not found:
            break
    
    if len(assigned) < required:
        missing = required - len(assigned)
        warning_log.append(
            f"{day}-{time_slot}: 需要{required}人，实际分配{len(assigned)}人 "
            f"(候选：{candidates_sorted})"
        )
    return len(assigned)

# 第一阶段分配
unassigned = set(all_people)
for day in ["周一", "周二", "周三", "周四", "周五"]:
    for time_slot in DAYTIME_SLOTS:
        assign_people(day, time_slot, TIME_SLOT_REQUIREMENT[time_slot])
    # 处理晚间时段
    time_slot = "10-11节"
    assigned_count = assign_people(day, time_slot, TIME_SLOT_REQUIREMENT[time_slot])
    if 0 < assigned_count < TIME_SLOT_REQUIREMENT[time_slot]:
        warning_log.append(f"{day}-{time_slot}: 晚间时段分配不足（{assigned_count}/{TIME_SLOT_REQUIREMENT[time_slot]}）")

# 第二阶段：强制分配剩余人员
while unassigned:
    person = next(iter(unassigned))
    assigned_flag = False
    
    # 优先尝试白天时段
    for day in ["周一", "周二", "周三", "周四", "周五"]:
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
        for day in ["周一", "周二", "周三", "周四", "周五"]:
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
    
    if not assigned_flag:
        warning_log.append(f"无法为 {person} 安排值班，请检查可用时间")
        unassigned.remove(person)

# 创建Excel写入对象
writer = pd.ExcelWriter('排班结果.xlsx', engine='openpyxl')
# 格式化输出
result = []
for day in ["周一", "周二", "周三", "周四", "周五"]:
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
result_df.T.to_excel(writer, sheet_name='排班表')
#创建人员统计表
stats_data = []
for person in sorted(all_people):
    stats = person_stats[person]
    days = ", ".join([d for d in stats["daily"] if stats["daily"][d] > 0])
    stats_data.append({
        "姓名": person,
        "总值班次数": stats["total"],
        "值班天数": len(stats["daily"]),
        "具体日期": days
    })
stats_df = pd.DataFrame(stats_data)
stats_df.to_excel(writer, sheet_name='人员统计', index=False)

# 创建警告日志表
if warning_log:
    warn_df = pd.DataFrame({"警告信息": warning_log})
    warn_df.to_excel(writer, sheet_name='异常提示', index=False)

# 设置单元格宽度
workbook = writer.book
for sheet_name in writer.sheets:
    worksheet = writer.sheets[sheet_name]
    if sheet_name == '排班表':
        worksheet.column_dimensions['A'].width = 12  # 时间列
        for col in ['B','C','D','E','F']:
            worksheet.column_dimensions[col].width = 25
    elif sheet_name == '人员统计':
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['D'].width = 30

# 保存文件
writer.close()

print("\n排班结果已保存至：排班结果.xlsx")
print("包含以下工作表：")
print("- 排班表：详细值班安排")
print("- 人员统计：个人值班汇总")
if warning_log:
    print("- 异常提示：需要人工处理的异常情况")

# 输出警告信息
'''if warning_log:
    print("\n重要警告：")
    for warn in warning_log:
        print("⚠️", warn)

# 验证结果
print("\n人员统计：")
for person in sorted(all_people):
    stats = person_stats[person]
    status = "✅" if stats['total'] >=1 else "❌"
    print(f"{status}{person}: 总次数={stats['total']}次, 每日分布=" + 
          ", ".join([f"{d}{stats['daily'][d]}次" for d in stats["daily"] if stats["daily"][d] > 0]))

print("\n时间段检查：")
for day in ["周一", "周二", "周三", "周四", "周五"]:
    for slot in schedule:
        count = len(assignments[day][slot])
        required = TIME_SLOT_REQUIREMENT[slot]
        # 晚间特殊处理
        if slot == "10-11节" and count >=1:
            continue
        if count < required:
            print(f"警告：{day}-{slot} 需要{required}人，实际{count}人")
'''