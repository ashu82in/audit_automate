#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun  2 17:13:00 2024

@author: ashutoshgoenka
"""

import pandas as pd
import numpy as np
import streamlit as st
from PIL import Image
import PIL
import os
import zipfile
from zipfile import ZipFile, ZIP_DEFLATED
import pathlib
import shutil
import docx
import docxtpl
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
from docx.shared import Cm, Inches
import random
from random import randint
from streamlit import session_state


st.set_page_config(layout="wide")

try:
    shutil.rmtree("images_comp_audit")
except:
    pass

try:
    os.mkdir("images_comp_audit")
except:
    pass

def createfile():
    document = Document()
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    document.save("Audit_Word.docx")


st.title("Upload Observation")
obs_file = st.file_uploader("Upload Observation Excel File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader", on_change=createfile)
if obs_file is not None:
    df = pd.read_excel(obs_file)
    df = df.dropna(thresh=5)
    df["obs_element"] = df["Element"] + "_"+df["Observations"]
    df["Location_Final"] =  df["Location"] + " at "+ df["Level"]
    df = df.sort_values("obs_element")
    df["Image Number"] =  df["Image Number"].astype(str)
    df["Image No Final"] = ""
    # df["Image List"] = ""
    df["Start Image Number"] = 0
    df["End Image Number"] = 0
    image_col = df.columns.get_loc("Image Number")
    st.write(image_col)
    st.write(df)
    start_val = 1
    end_val = 1
    img_master_list = []
    img_new_old_dict = {}
    for idx, row in df.iterrows():
        st.write(idx)
        st.write(row[image_col])
        st.write(df.loc[1,"Observations"])
        temp_value = row[image_col].split(",")
        temp_value = [t.strip() for t in temp_value]
        st.write(temp_value)
        img_master_list = img_master_list + temp_value
        no_of_img = int(len(temp_value))
        # df.loc[idx, "Image List"] = temp_value
        df.loc[idx, "No of Images"] = no_of_img
        df.loc[idx, "Start Image Number"]  = start_val
        df.loc[idx, "End Image Number"]  = start_val+no_of_img -1
        if no_of_img>1:
            df.loc[idx, "Image No Final"] = "Image "+str(start_val) +" - " + str(start_val+no_of_img -1)
        else:
            df.loc[idx, "Image No Final"] = "Image "+str(start_val)
        for ctr in range(no_of_img):
            img_new_old_dict[temp_value[ctr]] = start_val + ctr
                
        start_val = start_val+no_of_img
        end_val = start_val
    st.write(df)
    st.write(img_master_list)
    st.write(img_new_old_dict)
    
    up_files = st.file_uploader("Upload Image Files", type = ["png", "jpeg", "jpg"] ,accept_multiple_files=True)
    if up_files is not None:
        file_name_list = [file.name for file in up_files]
        file_not_found = []
        file_found = []
        image_file_dict_final = {}
        file_image_dict= {}
        for img_name in img_master_list:
            temp_found = [s for s in file_name_list if img_name in s]
            if len(temp_found)>0:
                file_name_list = file_name_list + temp_found
                image_file_dict_final[img_new_old_dict[img_name]] = temp_found[0]
                file_image_dict[temp_found[0]] = img_new_old_dict[img_name]
            else:
                file_not_found.append(img_name)
        
        st.write(image_file_dict_final)
        st.write(file_not_found)
        st.write(file_found)
        st.write(file_image_dict)
        for temp_file in up_files:
            try:
                st.write(temp_file.name)
                # st.write(file_image_dict[temp_file.name])
                img_no = file_image_dict[temp_file.name] 
                st.write(img_no)
                ext= temp_file.name.split(".")[-1]
                im = Image.open(temp_file)
                im.save("images_comp_audit/Image "+str(img_no)+"."+ext, exif=im.info.get("exif"))
                
            except:
                pass
        zip_path = "images_compressed_audit.zip"
        directory_to_zip = "images_comp_audit"
        folder = pathlib.Path(directory_to_zip)
        with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zip:
            for file in folder.iterdir():
                zip.write(file, arcname=file.name)
                
        with open("images_compressed_audit.zip", "rb") as fp:
            btn = st.download_button(
                label="Download ZIP",
                data=fp,
                file_name="images_compressed_audit.zip",
                mime="application/zip"
            )
        
            
            
            
            
            
        
        
# st.write(up_files)
    
    