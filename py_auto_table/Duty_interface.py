import streamlit as st
from utilics import runUtilics
import time

#video_file = open("py_auto_table\videos\web1.mp4","rb")
#video_bytes = video_file.read()
curtime = time.localtime()

st.title(f"{curtime[0]}学年招就处学生助理排班系统")

st.video("py_auto_table/videos/entireweb2.mp4",start_time=0,autoplay=True,loop=True)

st.divider()

st.subheader("🎈设置各班人员数🎈")

columns1,columns2 = st.columns(2)
#设置每个时间段值班人数
with columns1:
    to_one = st.slider("早上第一班(1-2)值班人数",min_value=1,max_value=3,step=1,value=2)
    to_two = st.slider("早上第二班(3-5)值班人数",min_value=2,max_value=4,step=1,value=3)
    nx_one = st.slider("下午第一班(6-7)值班人数",min_value=1,max_value=3,step=1,value=2)
with columns2:
    nx_two = st.slider("下午第二班(8-9)值班人数",min_value=2,max_value=4,step=1,value=3)
    night = st.slider("晚上值班(10-11)人数",min_value=0,max_value=2,step=1,value=2)
st.divider()
#获取上传的文件
upload_file = st.file_uploader("请上传人员数据表(200MB以内),仅支持excel格式",type=["xlsx","csv"])
#将三个表格置于不同的tab中
tableofworker,tableofstatics,tableofwarning = st.tabs(["值班助理工作表","值班助理统计数据表","排班警告提示表"])

#判断是否上传文件
if upload_file:
    with st.spinner("正在加载中...."):
        time.sleep(2)
        df_worker,df_statics,df_warning = runUtilics(upload_file,to_one,to_two,nx_one,nx_two,night)
    st.balloons()
#显示值班表数据
with tableofworker:
    if upload_file:
        st.write("程序生成值班表")
        #这里的.T使得原本数据表格中，行和列互换了
        st.table(df_worker.T)
        st.divider()
        st.write("待修改值班表(若无任何警告，可无需修改，直接下载即可)")
        st.data_editor(df_worker.T)
    else:
        st.write("内容为空，请上传文件！")
with tableofstatics:
    if upload_file:
        st.dataframe(df_statics,hide_index=True)
    else:
        st.write("内容为空，请上传文件！")
with tableofwarning:
    if upload_file:
        st.dataframe(df_warning,hide_index=True)
    else:
        st.write("内容为空，请上传文件！")

   
    

