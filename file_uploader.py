import streamlit as st
import streamlit.components.v1 as components
import base64
import json
from io import BytesIO

def custom_file_uploader(label="上传文件", key="customUploader"):
    uploaded_file = None
    uploaded_name = None

    uploaded_json = components.html(f"""
    <input type="file" id="fileInput" />
    <script>
        const input = document.getElementById("fileInput");
        input.addEventListener("change", function() {{
            const file = input.files[0];
            const reader = new FileReader();
            reader.onload = function() {{
                const base64 = reader.result.split(',')[1];
                const payload = {{
                    filename: file.name,
                    content: base64
                }};
                window.parent.postMessage({{ type: 'streamlit:setComponentValue', value: payload }}, '*');
            }};
            reader.readAsDataURL(file);
        }});
    </script>
    """, height=100, key=key)

    # 接收 base64 数据
    if isinstance(uploaded_json, dict) and "content" in uploaded_json and "filename" in uploaded_json:
        decoded_bytes = base64.b64decode(uploaded_json["content"])
        uploaded_file = BytesIO(decoded_bytes)
        uploaded_name = uploaded_json["filename"]

    return uploaded_name, uploaded_file
