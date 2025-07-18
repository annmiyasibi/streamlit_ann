import streamlit as st
from PIL import Image 

st.title(" Simple Streamlit App")

st.write("Welcome to this simple Streamlit app! ")


image = Image.open("C:/Users/annmi/Downloads/62.jpg")


st.image(
    image,
    caption="Sample Image",
    use_column_width=True
)

st.success("This is a success message!")
