import pandas as pd
import streamlit as st
st.title("upload csv")
st.write("choose the csv file required")
uploaded_file=st.file_uploader("browse files")
if uploaded_file:
  file=pd.read_csv(uploaded_file)
  st.write(file)
  




