import streamlit as st
import time
from sentMail import feedBack

with st.sidebar:
    st.subheader("About me")
    st.text("作者：(•̀ᴗ•́)و ̑̑小王先生")
    st.text("鸣谢：江西财经大学招生就业处!")
    st.divider()
    st.text("江财百年校庆——《百载》")
    st.audio("py_auto_table/sounds/background.mp4",autoplay=False,start_time=0,loop=True)
    st.text("江财校歌——《信敏廉毅》")
    st.audio("py_auto_table/sounds/background2.mp4",autoplay=True,start_time=0,loop=True)
    st.text("江财rap——《"'赣'"出未来》")
    st.audio("py_auto_table/sounds/background3.mp4",autoplay=False,start_time=0,loop=True)

st.title("(*^‧^*)网站使用说明")
st.divider()
st.subheader("( •̅_•̅ )界面一：学生助理自动排班系统")
st.write("( •̅_•̅ )为了方便数据处理，导入文件内容格式尽量保持如下形式:")
st.image("py_auto_table/images/showWorker.png")
st.divider()
note1,note2,note3 = st.tabs(["注意点一","注意点二","注意点三"])
with note1:
    st.write("( •̅_•̅ )时间列必须是1-2节和星期几的形式,不可加其他文字;\n助理姓名需要使用英文逗号隔开(用中文应该也行,但是我没有试)")
with note2:
    st.write("( •̅_•̅ )程序生成可能没这么完美，部分位置还需要动手调！")
with note3:
    st.write("🎈遇到问题根据后面联系方式进行反馈,网站也会不断更新！")

st.subheader("( •̅_•̅ )界面二：企业对接系统")
st.write("( •̅_•̅ )为了方便数据处理，导入文件内容格式尽量保持如下形式:")
st.image("py_auto_table/images/showEnterprise.png")
st.divider()
E_note1,E_note2,E_note3 = st.tabs(["注意点一","注意点二","注意点三"])
with E_note1:
    st.write("( •̅_•̅ )可以直接使用企业源数据表，具体样式，问老助理！")
with E_note2:
    st.image("py_auto_table/images/webWorker.png")
    st.write("( •̅_•̅ )值班助理表格直接使用此网站页面一，下载的样式即可，具体样式如上图")
with E_note3:
    st.write("🎈遇到问题根据后面联系方式进行反馈,网站也会不断更新！")
st.divider()
st.subheader("(๑・ˍ・)联系我")
st.divider()
info_name = st.text_input("(๑・ˍ・)请输入您的姓名，以便更好的了解信息来源!")
text = st.text_area("(๑・ˍ・)请输入你想要反馈的内容")
submit = st.button("发送")
if submit:
    if not info_name:
        st.warning("请填写上方姓名框")
    if not text:
        st.warning("请输入您需要反馈的内容!")
    else:
        with st.spinner("发送中...."):
            time.sleep(2)
            st.success("提交成功！等待回复")
            st.balloons()
        isSucess = feedBack(info_name,text)
        if not isSucess:
            st.warning("发送失败，注意！需要使用QQ邮箱")