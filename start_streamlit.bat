@echo off
cd /d C:\Users\Administrator\Desktop\codewhale
start http://localhost:8501
streamlit run translator_streamlit.py --server.port 8501
