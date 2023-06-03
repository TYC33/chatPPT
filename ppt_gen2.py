import streamlit as st
from ppt_lib import *
import base64
import re
import openai

def create_download_link(data, filename):
    b64 = base64.b64encode(data).decode()  # 將檔案數據轉換為 base64 編碼
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">點此下載</a>'
    #建立超連結
    return href

def split_list(input_list, chunk_size):    #看有幾個小點做分頁
    return [input_list[i:i + chunk_size] for i in range(0, len(input_list), chunk_size)]
    
def remove_punctuation(text):
    # 使用正則表達式移除標點符號
    cleaned_text = re.sub(r'[^\w\s]', '', text)
    return cleaned_text

if __name__ == '__main__':
    if 'context' not in st.session_state:   #若內文不在st.session_state
        st.session_state.context = ''       #則為空

    st.set_page_config(layout = 'wide')     #增加寬度
    
    st.title("ChatPPT")                     #標題
    
    col6, col7 = st.columns(2)              #分為兩個相同大小的column
                                            #一個為title的，另一個為sub_title
    title = col6.text_input("請輸入標題:", value = '要如何學習機器學習')
    sub_title = col7.text_input("請輸入副標題:", value = '陳姿羽/銘傳大學資工系')
    
    col8, col9 = st.columns([1,4])          #分為兩個大小不同的column
                                            #context_title占比20%, context占比80%
    context_title = col8.text_input("請輸入內文標題:", value = '學習機器學習的步驟')
    context = col9.text_area("請輸入內文:", value = st.session_state.context, height=170)
    
    col1, col2, col3, col4, col5 = st.columns(5)  #分為五個大小相同的column
    #chatGPT = col1.checkbox('ChatGPT導入內文', value=False)  #預設值為不勾選
    #topic = col2.text_input("您希望ChatGPT是何種專家:", value = '機器學習')
    #temperature = col3.slider('ChatGPT的創意程度', 0.0, 1.0, 0.0, 0.1)
    num_of_points = col5.number_input("投影片每頁要列出幾點:", value = 3)
    
    if st.button("產生PPT"):
        
        # 產生投影片
        prs = Presentation()  #建立投影片
        
        sub_title = sub_title.replace('/', '\n') #去除副標題的"/"與換行
        
        create_title(prs, title, sub_title)      #建立新的投影片標題
        
        if len(context) > 0:
            item_list = []
            
            for line in context.splitlines():    #以一行一行讀取每個context
                item_list.append(line.strip())   #去除換行並加入item_list
            
            chunked_lists = split_list(item_list, num_of_points) #看有幾個小點做分頁

            slides = []

            # Print the chunked lists
            for sublist in chunked_lists: #加入投影片標題與內容
                slides.append({'title': context_title, 'content': sublist})
            
            create_body(prs, slides)
    
            filename = remove_punctuation(title) + '.pptx'  #存取投影片的名稱
            prs.save(filename)   

            with open(filename, 'rb') as f: #以二進位方式讀檔
                data = f.read()

            #顯示下載連結
            st.markdown(create_download_link(data, filename), unsafe_allow_html=True)