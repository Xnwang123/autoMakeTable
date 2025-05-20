import pandas as pd
import re

def runMakeEnterprise():
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
    # 主逻辑
    if __name__ == "__main__":
        #生成初版企业对接表
        extracted_df = makeFile("py_auto_table\sourcedata.xlsx")
        # 加载课程表数据
        schedule_df = load_schedule_data('py_auto_table/2025-05-19T09-16_export.csv')
        
        # 从生成的初版企业对接表中读取已有的宣讲会数据
        extracted_df = makeFile("py_auto_table\sourcedata.xlsx")
        
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
        # 保存更新后的数据
        #extracted_df.to_excel('py_auto_table/extracted_data.xlsx', index=False, sheet_name='企业对接表')
        #print("对接助理已成功添加至 extracted_data.xlsx！")

extracted_df = runMakeEnterprise()
extracted_df.to_excel('py_auto_table/extracted_data.xlsx', index=False, sheet_name='企业对接表')
print("对接助理已成功添加至 extracted_data.xlsx！")