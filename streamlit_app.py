import streamlit as st 
import pandas as pd
from io import StringIO

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE

from pptx.util import Cm
from pptx.dml.color import RGBColor
from PIL import Image
import os
import copy


def process_slide(slide_src):
    # 获取页面中的所有形状
    shapes = slide_src.shapes

    # 遍历所有形状
    for shape in shapes:
        # 判断形状类型是否为图片
        print(shape.name)
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # 处理图片
            #process_image(shape, slide_dst)
            if shape.width >Cm(0.1) and shape.height > Cm(0.1):
                slide_dst = prs_dst.slides.add_slide(blank_slide_layout)

                imdata = shape.image.blob
                imagetype = shape.image.content_type
                typekey = imagetype.find('/') + 1
                imtype = imagetype[typekey:]
                image_file ="tmp."+imtype
                file_str=open(image_file,'wb')
                file_str.write(imdata)
                file_str.close()

                new_shape_width=shape.width*3
                new_shape_height=shape.height*3
                new_shape_top=prs_dst.slide_width/4-new_shape_height/2
                new_shape_left=prs_dst.slide_width/4-new_shape_width/2
                
                new_shape = slide_dst.shapes.add_picture(image_file,new_shape_top,new_shape_left,new_shape_width,new_shape_height)

                gap_width=0
                gap_height=0

                txBox = slide_dst.shapes.add_textbox(prs_dst.slide_width/2,0,prs_dst.slide_width/2,prs_dst.slide_height)
                tf = txBox.text_frame
                tf.word_wrap = False
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                #tf.fit_text(font_family='Calibri', max_size=72, bold=False, italic=False, font_file=None)

                print("tf lenth:",len(tf.text))
                for shape_text in shapes:
                    print("----",shape_text.name,"--",shape_text.shape_type)
                    #if shape_text.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                    if shape_text.has_text_frame:
                        print("---- start to copy content")
                        if len(tf.text)==0:
                            gap_width=abs(shape_text.left - shape.left)
                            gap_height=abs(shape_text.top-shape.top)
                            tf.text = shape_text.text_frame.text
                            print("----this is the first time",shape_text.text_frame.text)
                        else:
                            if gap_width>abs(shape_text.left - shape.left) and gap_height>abs(shape_text.top-shape.top):
                                print("----find closer one",shape_text.text_frame.text)
                                gap_width=abs(shape_text.left - shape.left)
                                gap_height=abs(shape_text.top-shape.top)
                                if len(shape_text.text_frame.text)>0:
                                    tf.text = shape_text.text_frame.text
                        

 
         
        # elif shape.shape_type == MSO_SHAPE_TYPE.LINE
        #     shapes.element.remove(shape.element)



def app_head():
    #列举当前功能信息
    st.markdown("铭铭的英语学习卡片转换助手")
    st.write("已经实现功能")
    asis_data = {
        "时间": 
            ["2024-03-26",
        ],
        "功能": 
            ["读取固定ppt源文件生成固定目标ppt"
            "将源文件中的图片分别生成对应的A5尺寸的ppt页面"
            "将相关的文本拷贝到图片左侧，形成卡片",
        ]
    }
    asis_df = pd.DataFrame(asis_data)
    st.write(asis_df)

    #列举待实现的功能信息
    st.write("规划中的功能")
    rdmap_data = {
        "时间": 
            ["2024-03-31",
            "2024-04-20",
            
        ],
        "功能": 
            ["文件上传下载"
            "文字匹配更加智能，同时支持格式自动生成",
            "加入对pdf的支持",
        ]
    }
    rdmap_df = pd.DataFrame(rdmap_data)
    st.write(rdmap_df)


app_head()

uploaded_files=st.file_uploader("请上传需要调整格式的ppt文件，仅支持ppt/pptx,可同时上传多个文件",accept_multiple_files=True)
#uploaded_files.type=['ppt','pptx']
for uploaded_file in uploaded_files:
    bytes_data = uploaded_file.read()
    with open("uploaded_file.pptx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.write("filename:", uploaded_file.name)
    #st.write(bytes_data)
    # 打开原始和目标 PPT
    prs_src = Presentation("uploaded_file.pptx")


    # 新建目标 PPT
    prs_dst = Presentation()
    prs_dst.slide_height=Cm(14.8)
    prs_dst.slide_width=Cm(21.0)
    blank_slide_layout=prs_dst.slide_layouts[6]
    # 遍历原始 PPT 中的每一页
    for slide_src in prs_src.slides:
        # 处理当前页面的图片
        process_slide(slide_src)

    # 保存目标 PPT
    try:
        prs_dst.save("target.pptx")
    except FileExistsError:
        # 如果目标文件已存在，则直接覆盖
        prs_dst.save("target.pptx")

    with open('target.pptx', 'rb') as ff:
        target_file = ff
        st.download_button('下载转换后的pptx', target_file.read(),file_name="target.pptx",mime="pptx") 

#binary_contents = b'target.pptx'
#with open('myfile.zip', 'rb') as f:
#   st.download_button('Download target', f, file_name='target.ppx')


st.write("Here we are at the end of getting started with streamlit! Happy Streamlit-ing! :balloon:")
