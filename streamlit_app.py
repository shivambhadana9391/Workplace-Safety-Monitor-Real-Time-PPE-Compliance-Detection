import streamlit as st

st.set_page_config(page_title="My First App", layout="centered")

st.title("Welcome to Streamlit App")

name = st.text_input("Enter your name")

if name:
    st.success(f"Hello, {name}!")

num1 = st.number_input("Enter first number", value=0)
num2 = st.number_input("Enter second number", value=0)

if st.button("Add"):
    st.write("Result:", num1 + num2)

option = st.selectbox("Choose a topic", ["Python", "C", "Java"])

if option == "Python":
    st.write("Python is easy and powerful.")
elif option == "C":
    st.write("C is fast and low-level.")
elif option == "Java":
    st.write("Java is object-oriented and widely used.")

uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    st.write("File uploaded:", uploaded_file.name)

st.slider("Rate this app", 1, 10)

st.write("Thank you for using this app!")
