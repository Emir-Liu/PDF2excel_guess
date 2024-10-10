import re
import sys
from typing import List

import pandas as pd
import pymupdf

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def get_order_number_single_page(page, order_number_table) -> str:
    """get order number in page

    Args:
        order_number_table (_type_): _description_

    Returns:
        str: _description_
    """

    # row = order_number_table.rows[0]
    # cell_order_number = row.cells[1]
    # order_number = page.get_text(option="text", clip=cell_order_number)
    # print(f"order num:{order_number}")
    order_number = get_cell_content_in_table(
        page=page, table=order_number_table, num_row=0, num_cell=1
    )
    # print(f"order num:{order_number}")
    return order_number


def get_cell_content_in_table(page, table, num_row, num_cell) -> str:
    row = table.rows[num_row]
    cell = row.cells[num_cell]
    content = page.get_text(option="text", clip=cell).strip()
    return content


def get_table_pos(table) -> (float, float):
    table_up_y = table.bbox[1]
    table_bo_y = table.bbox[3]
    return table_up_y, table_bo_y


def get_style_pos_y_info_list(page, tables, page_height) -> List[dict]:
    """get pos of styles in single page

    Args:
        page (_type_): _description_
        tables (_type_): _description_

    Returns:
        List[dict]: _description_
        [
            {
                "line_bot_y": -1,
                "comp_up_y": -1,
                "comp_bot_y": -1,
                "hts_up_y": -1,
                "hts_bot_y": -1,
                "end_style_y": -1,
            },
            ...
        ]
    """

    # order_number = get_order_number_single_page(page=page, order_number_table=tables[0])
    mid_pos_info_list = []
    mid_pos_info = {
        "line_bot_y": -1,
        "comp_up_y": -1,
        "comp_bot_y": -1,
        "hts_up_y": -1,
        "hts_bot_y": -1,
        "end_style_y": -1,
    }
    num_table = 0
    for table in tables:
        num_table += 1
    for table_idx, table in enumerate(tables):
        content = get_cell_content_in_table(
            page=page, table=table, num_row=0, num_cell=0
        )

        num_table_rows = len(table.rows)

        # print(f"content:{content}")
        if "Order Number" in content and num_table_rows == 1:
            # first table:
            # order number
            continue
        elif table_idx == num_table - 1:
            # last table:
            # Total Order or HTS
            if "Total Order:" in content:
                # print("end")
                # mid_pos_info["hts_up_y"], mid_pos_info["hts_bot_y"] = get_table_pos(
                #     table=table
                # )
                mid_pos_info["end_style_y"], _ = get_table_pos(table=table)
                mid_pos_info_list.append(mid_pos_info)
                break
            elif "HTS" in content:
                mid_pos_info["hts_up_y"], mid_pos_info["hts_bot_y"] = get_table_pos(
                    table=table
                )
                mid_pos_info["end_style_y"] = page_height
                mid_pos_info_list.append(mid_pos_info)
                mid_pos_info = {
                    "line_bot_y": -1,
                    "comp_up_y": -1,
                    "comp_bot_y": -1,
                    "hts_up_y": -1,
                    "hts_bot_y": -1,
                    "end_style_y": -1,
                }
            # manage the previous info
        else:
            # mid table:
            # Line, Composition, HTS
            # get first cell in first row of the table

            if "Line" in content or "Order Number" in content and num_table_rows == 2:
                if mid_pos_info["line_bot_y"] == -1:
                    # first style info
                    pass
                else:
                    # manage the previous info
                    mid_pos_info["end_style_y"], _ = get_table_pos(table=table)
                    # print(f"mid_pos_info:{mid_pos_info}")
                    mid_pos_info_list.append(mid_pos_info)
                    mid_pos_info = {
                        "line_bot_y": -1,
                        "comp_up_y": -1,
                        "comp_bot_y": -1,
                        "hts_up_y": -1,
                        "hts_bot_y": -1,
                        "end_style_y": -1,
                    }

                    # pass
                _, mid_pos_info["line_bot_y"] = get_table_pos(table=table)
            elif "Composition" in content:
                mid_pos_info["comp_up_y"], mid_pos_info["comp_bot_y"] = get_table_pos(
                    table=table
                )
            elif "HTS" in content:
                mid_pos_info["hts_up_y"], mid_pos_info["hts_bot_y"] = get_table_pos(
                    table=table
                )

            else:
                print(f"格式解析出现意外:{content}")
        # print(f"mid_pos_info:{mid_pos_info}")

    # print(f"mid_pos_info_list:{mid_pos_info_list}")
    return mid_pos_info_list


def get_row_content(page, up_y, down_y, page_width) -> List[str]:
    content = page.get_text(
        option="dict",
        clip=(
            0,
            up_y,
            page_width,
            down_y,
        ),
    )

    # print(f"content:{content}")
    span_list = []

    for block in content["blocks"]:
        if "lines" in block:
            for line in block["lines"]:

                for span in line["spans"]:
                    span_content = span["text"].strip()
                    # print(f"span_content:{span_content}")
                    span_list.append(span_content)
    return span_list


def get_row_content_block(page, up_y, down_y, page_width) -> List[str]:
    content = page.get_text(
        option="dict",
        clip=(
            0,
            up_y,
            page_width,
            down_y,
        ),
    )

    # print(f"content:{content}")
    block_list = []

    for block in content["blocks"]:
        tmp_block_list = []
        if "lines" in block:
            for line in block["lines"]:

                for span in line["spans"]:
                    span_content = span["text"].strip()
                    # print(f"span_content:{span_content}")
                    tmp_block_list.append(span_content)

        tmp_block_content = " ".join(tmp_block_list)
        block_list.append(tmp_block_content)
    return block_list


def get_style_content(mid_pos_info_list, order_number, page, page_width) -> List[dict]:

    style_info_list = []
    for tmp_mid_pos_info in mid_pos_info_list:
        tmp_style_info = {
            "款号": "",
            "PO": order_number,
            "色号": "",
            "尺寸": [],
            "总数": -1,
            "交期": [],
            "价格": "",
        }

        span_list = get_row_content(
            page=page,
            up_y=tmp_mid_pos_info["line_bot_y"],
            down_y=tmp_mid_pos_info["comp_up_y"],
            page_width=page_width,
        )

        # print(f"span_list:{span_list}")
        tmp_style_info["款号"] = span_list[1]
        tmp_style_info["色号"] = span_list[2]
        tmp_style_info["交期"] = str([span_list[-2], span_list[-1]])

        span_list = get_row_content(
            page=page,
            up_y=tmp_mid_pos_info["comp_bot_y"],
            down_y=tmp_mid_pos_info["hts_up_y"],
            page_width=page_width,
        )
        # print(f"span_list:{span_list}")
        tmp_style_info["价格"] = span_list[0]

        block_list = get_row_content_block(
            page=page,
            up_y=tmp_mid_pos_info["hts_bot_y"],
            down_y=tmp_mid_pos_info["end_style_y"],
            page_width=page_width,
        )

        size_info_list = []
        bool_start = False
        for tmp_block_content in block_list:
            if "Total" in tmp_block_content:
                bool_start = not bool_start
            tmp_size_info = tmp_block_content.split(" ")
            if bool_start is True:
                size_info_list.append(tmp_size_info)
            if len(size_info_list) and bool_start is False:
                size_info_list.append(tmp_size_info)
                break

        # print(f"size_info_list:{size_info_list}")

        tmp_style_info["总数"] = size_info_list[-1][1]
        tmp_style_info["尺寸"] = size_info_list

        len_first_row = len(size_info_list[0])
        # print(f"len_first_row:{len_first_row}")

        first_row = size_info_list[0]
        for row_idx, row in enumerate(size_info_list):
            # print(f"row_idx:{row_idx} row:{row}")
            if row_idx == 0:
                continue
            elif row_idx == len(size_info_list) - 1:
                continue
            if len(row) > len_first_row:
                # include 内裤长
                new_size = row[0]
                tmp_style_info["色号"] += f"--内长{new_size}英寸"
                start_pos = 2
                # size = [fr + f"/{new_size}" for fr in first_row[1:]]
                size = first_row[1:]
            else:
                start_pos = 1
                size = first_row[1:]
            for tmp_size_idx, tmp_size in enumerate(size):
                tmp_style_info[tmp_size] = row[start_pos + tmp_size_idx]
            # tmp_style_info[""]

        # print(f"tmp_style_info:{tmp_style_info}")
        style_info_list.append(tmp_style_info)
        # get style code and others
    return style_info_list


def clean_annot_in_doc(doc):
    # remove annotation information from pdf files
    # to avoid the impact of annotation information on form extraction
    for page in doc:
        for annot in page.annots():
            page.delete_annot(annot=annot)


def func_pdf2excel(pdf_content):
    # convert reading local file into reading data stream,
    # avoiding the need to save the file locally
    doc = pymupdf.open(stream=pdf_content)
    clean_annot_in_doc(doc=doc)
    # get PO
    page = doc[0]
    page_width = page.rect[2]
    tables = page.find_tables()
    order_number = get_order_number_single_page(page=page, order_number_table=tables[0])

    # get target country
    _, up_target_country = get_table_pos(table=tables[0])
    down_target_country, _ = get_table_pos(table=tables[1])

    content = page.get_text(
        option="text", clip=(0, up_target_country, page_width, down_target_country)
    )

    print(f"content:{content}")

    pattern = r"\((.*?)\)"

    match_obj = re.search(pattern, content)
    target_country = ""
    if match_obj:
        target_country = match_obj.group(1)
        print(f"target_country:{target_country}")
    # sys.exit()
    order_number += target_country
    total_style_info_list = []
    for page in doc:
        # find struct by tables
        # page = doc[0]
        tables = page.find_tables()
        num_table = 0
        for table in tables:
            num_table += 1
        page_width = page.rect[2]
        page_height = page.rect[3]

        mid_pos_info_list = get_style_pos_y_info_list(
            page=page, tables=tables, page_height=page_height
        )
        print(f"len mid_pos_info_list:{len(mid_pos_info_list)}")

        style_info_list = get_style_content(
            mid_pos_info_list=mid_pos_info_list,
            order_number=order_number,
            page=page,
            page_width=page_width,
        )
        total_style_info_list.extend(style_info_list)

    # translate json to excel file
    df = pd.DataFrame(total_style_info_list)
    print(f"df:{df}")
    df = df.drop("尺寸", axis=1)
    df["总数"] = df["总数"].astype(int)

    df = df.sort_index(axis=1)
    front_list = ["款号", "PO", "色号"]
    end_list = ["总数", "价格", "交期"]
    new_order = (
        front_list
        + [col for col in df.columns if col not in front_list + end_list]
        + end_list
    )
    exist_index = df.columns
    # print(f"exist_index:{exist_index}")
    for ordered_key in new_order:
        if ordered_key in exist_index:
            pass
        else:
            df[ordered_key] = pd.NA
    df = df[new_order]
    print(f"df:{df}")

    # df.to_excel("c.xlsx", index=False)
    # 创建一个Workbook对象
    wb = Workbook()

    # 获取当前活跃的工作表
    ws = wb.active

    # 将DataFrame的数据写入工作表
    for r in dataframe_to_rows(df, index=False, header=True):
        print(f"r:{r}")
        ws.append(r)

    # wb.save("back.xlsx")
    return wb


if __name__ == "__main__":

    # ORG_PDF_PATH = "D:/projects/pdf2excel/pdf2excel_GUESS/others/sample_file/Order_EU_03_MA01-2024-02483_202406060132090609.pdf"
    ORG_PDF_PATH = "D:/projects/pdf2excel/pdf2excel_GUESS/others/sample_file/Order_EU_03_AU01-2024-00882_202406060143288595.pdf"
    doc = pymupdf.open(ORG_PDF_PATH)

    # get PO
    page = doc[0]
    page_width = page.rect[2]
    tables = page.find_tables()
    order_number = get_order_number_single_page(page=page, order_number_table=tables[0])

    # get target country
    _, up_target_country = get_table_pos(table=tables[0])
    down_target_country, _ = get_table_pos(table=tables[1])

    content = page.get_text(
        option="text", clip=(0, up_target_country, page_width, down_target_country)
    )

    print(f"content:{content}")

    pattern = r"\((.*?)\)"

    match_obj = re.search(pattern, content)
    target_country = ""
    if match_obj:
        target_country = match_obj.group(1)
        print(f"target_country:{target_country}")
    # sys.exit()
    order_number += target_country
    total_style_info_list = []
    for page in doc:
        # find struct by tables
        # page = doc[0]
        tables = page.find_tables()
        num_table = 0
        for table in tables:
            num_table += 1
        page_width = page.rect[2]
        page_height = page.rect[3]

        mid_pos_info_list = get_style_pos_y_info_list(
            page=page, tables=tables, page_height=page_height
        )
        print(f"len mid_pos_info_list:{len(mid_pos_info_list)}")

        style_info_list = get_style_content(
            mid_pos_info_list=mid_pos_info_list,
            order_number=order_number,
            page=page,
            page_width=page_width,
        )
        total_style_info_list.extend(style_info_list)

    # translate json to excel file
    df = pd.DataFrame(total_style_info_list)

    df = df.drop("尺寸", axis=1)
    df["总数"] = df["总数"].astype(int)

    df = df.sort_index(axis=1)
    front_list = ["款号", "PO", "色号"]
    end_list = ["总数", "价格", "交期"]
    new_order = (
        front_list
        + [col for col in df.columns if col not in front_list + end_list]
        + end_list
    )
    exist_index = df.columns
    # print(f"exist_index:{exist_index}")
    for ordered_key in new_order:
        if ordered_key in exist_index:
            pass
        else:
            df[ordered_key] = pd.NA
    df = df[new_order]
    print(f"df:{df}")

    df.to_excel("c.xlsx", index=False)
