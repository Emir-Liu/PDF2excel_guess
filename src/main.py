import os
import sys
import io
from io import BytesIO
from typing import List


ROOT_PATH = os.path.dirname(os.path.abspath(__file__))
# print(f"ROOT_PATH:{ROOT_PATH}")
sys.path.append(ROOT_PATH)

import fastapi
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import uvicorn
import openpyxl

from configs import TITLE, VERSION, IP
from function.pdf2excel import func_pdf2excel, trans_json2ws

app = FastAPI(title=TITLE, version=VERSION)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/uploadpdf/")
async def upload_pdf(file: List[UploadFile] = File(...)):
    total_style_info_list = []
    size_columns_set = set()
    for tmp_file in file:
        contents = await tmp_file.read()
        # file_name = file.filename
        # new_file_name = file_name.replace(" ", "_")
        bytes_io = io.BytesIO(contents)
        part_style_info_list, size_columns_set = func_pdf2excel(
            pdf_content=bytes_io, size_columns_set=size_columns_set
        )
        total_style_info_list.extend(part_style_info_list)

    excel_content = trans_json2ws(
        total_style_info_list, size_columns_set=size_columns_set
    )

    # 将 Excel 文件保存到 BytesIO 对象中
    output = BytesIO()
    excel_content.save(output)
    output.seek(0)  # 重置文件指针到开始位置

    # 创建一个 StreamingResponse 对象
    headers = {"Content-Disposition": f'attachment; filename="sheet.xlsx"'}
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


if __name__ == "__main__":
    uvicorn.run(
        app,
        host=IP,
        port=12300,
        # host=server_ip,
        # port=server_port
    )
