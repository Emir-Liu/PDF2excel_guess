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


def get_style_content(
    mid_pos_info_list, order_number, page, page_width, size_columns_set
) -> List[dict]:

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
            # tmp_size_info_int = [int(val) for val in tmp_size_info]
            tmp_size_info_int = []

            for val in tmp_size_info:
                try:
                    tmp_val = int(val)
                except Exception as e:
                    tmp_val = val
                tmp_size_info_int.append(tmp_val)
            if bool_start is True:
                size_info_list.append(tmp_size_info_int)
            if len(size_info_list) and bool_start is False:
                size_info_list.append(tmp_size_info_int)
                break

        print(f"size_info_list:{size_info_list}")

        tmp_style_info["总数"] = size_info_list[-1][1]
        tmp_style_info["尺寸"] = size_info_list

        tmp_size_columns = size_info_list[0][1:]
        for tmp_tmp_size_columns in tmp_size_columns:
            size_columns_set.add(tmp_tmp_size_columns)

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
    return style_info_list, size_columns_set


def clean_annot_in_doc(doc):
    # remove annotation information from pdf files
    # to avoid the impact of annotation information on form extraction
    for page in doc:
        for annot in page.annots():
            page.delete_annot(annot=annot)


def sort_size_list(size_set):
    size_to_number = {
        "XXXS": 1,
        "XXS": 2,
        "XS": 3,
        "S": 4,
        "M": 5,
        "L": 6,
        "XL": 7,
        "XXL": 8,
        "XXXL": 9,
    }

    def sort_sizes_str(sizes):
        sorted_sizes = sorted(
            sizes, key=lambda size: size_to_number.get(size, 0), reverse=False
        )
        return sorted_sizes

    bool_all_number = True
    size_list = []
    for tmp_size_set in size_set:
        try:
            # tmp_size_set = int(tmp_size_set)
            size_list.append(int(tmp_size_set))
        except Exception as e:
            size_list.append(tmp_size_set)
            bool_all_number = False

    # for tmp_size in size_list:
    #     print(f"tmp_size:{tmp_size}")
    # print(f"bool_all_number:{bool_all_number}")
    if bool_all_number is False:
        sorted_size = sort_sizes_str(size_list)
    else:
        sorted_size = sorted(size_list, reverse=False)

    return sorted_size


def trans_json2ws(total_style_info_list, size_columns_set):

    # translate json to excel file
    df = pd.DataFrame(total_style_info_list)
    # print(f"df:{df}")
    df = df.drop("尺寸", axis=1)
    df["总数"] = df["总数"].astype(int)

    # df = df.sort_index(axis=1)
    front_list = ["款号", "PO", "色号"]
    end_list = ["总数", "价格", "交期"]

    size_list = sort_size_list(size_set=size_columns_set)
    new_order = front_list + size_list + end_list
    exist_index = df.columns

    for ordered_key in new_order:
        if ordered_key in exist_index:
            pass
        else:
            df[ordered_key] = pd.NA
    df = df[new_order]
    # print(f"df:{df}")

    # df.to_excel("c.xlsx", index=False)
    # 创建一个Workbook对象
    wb = Workbook()

    # 获取当前活跃的工作表
    ws = wb.active

    # 将DataFrame的数据写入工作表
    for r in dataframe_to_rows(df, index=False, header=True):
        print(f"r:{r}")
        ws.append(r)

    return wb


def get_cll(page, page_width, down_cll):

    cll = ""
    tmp_cll_contents = page.get_text(option="dict", clip=(0, 0, page_width, down_cll))

    for block in tmp_cll_contents["blocks"]:
        tmp_block_content = ""
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    tmp_block_content += span["text"].strip()
        if len(tmp_block_content) == 3:
            cll = tmp_block_content
            break

    return cll


def get_target_country(page, up_target_country, down_target_country, page_width):

    content = page.get_text(
        option="dict", clip=(0, up_target_country, page_width, down_target_country)
    )

    target_block_json = {"left": 0, "bot": 0, "content": ""}
    for block in content["blocks"]:
        # print(f"block:{block}")
        tmp_block_list = []
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    tmp_block_list.append(span["text"].strip())

        # block_info_json["content"] = " ".join(tmp_block_list)
        block_info_json = {
            "left": block["bbox"][0],
            "bot": block["bbox"][3],
            "content": " ".join(tmp_block_list),
        }

        # print(f"block info json:{block_info_json}")

        if block_info_json["left"] > page_width / 4:
            continue
        # tmp_block_content = ''
        if block_info_json["bot"] < target_block_json["bot"]:
            continue

        if "Payment" in block_info_json["content"]:
            break
        target_block_json = block_info_json
        print(f"new target_block_json:{target_block_json}")

    if target_block_json["content"]:
        target_block_json["content"] = target_block_json["content"].replace("(", " ")
        target_block_json["content"] = target_block_json["content"].replace(")", " ")
        target_country = target_block_json["content"].strip().split(" ")[-1]
    else:
        target_country = " "

    return target_country


def func_pdf2excel(pdf_content, size_columns_set):
    # convert reading local file into reading data stream,
    # avoiding the need to save the file locally

    # Distinguish whether data is a stream or a path string
    if isinstance(pdf_content, str):
        doc = pymupdf.open(pdf_content)
    else:
        doc = pymupdf.open(stream=pdf_content)

    clean_annot_in_doc(doc=doc)

    # get PO
    page = doc[0]
    page_width = page.rect[2]
    tables = page.find_tables()
    order_number = get_order_number_single_page(page=page, order_number_table=tables[0])

    # get CLL
    down_cll, up_target_country = get_table_pos(table=tables[0])
    cll = get_cll(page, page_width=page_width, down_cll=down_cll)

    # get target country
    down_target_country, _ = get_table_pos(table=tables[1])
    target_country = get_target_country(
        page, up_target_country, down_target_country, page_width
    )

    order_number = f"{order_number.strip()}-{target_country.strip()}-{cll.strip()}"
    total_style_info_list = []

    for page in doc:
        # find struct by tables
        tables = page.find_tables()
        num_table = 0
        for table in tables:
            num_table += 1
        page_width = page.rect[2]
        page_height = page.rect[3]

        mid_pos_info_list = get_style_pos_y_info_list(
            page=page, tables=tables, page_height=page_height
        )
        # print(f"len mid_pos_info_list:{len(mid_pos_info_list)}")

        style_info_list, size_columns_set = get_style_content(
            mid_pos_info_list=mid_pos_info_list,
            order_number=order_number,
            page=page,
            page_width=page_width,
            size_columns_set=size_columns_set,
        )
        total_style_info_list.extend(style_info_list)

    # translate json to excel file
    return total_style_info_list, size_columns_set


if __name__ == "__main__":
    # ORG_PDF_PATH = "D:/projects/pdf2excel/pdf2excel_GUESS/others/sample_file/Order_EU_03_MA01-2024-02483_202406060132090609.pdf"
    # ORG_PDF_PATH = "D:/projects/pdf2excel/pdf2excel_GUESS/others/sample_file/Order_EU_03_MA03-2024-03573_202409061447208210.pdf"
    ORG_PDF_PATH = "C:/Users/liuyiming/Downloads/OCR1010/error/Order_EU_03_MA01-2024-04031_202409061439279596.pdf"
    # ORG_PDF_PATH = "C:/Users/liuyiming/Downloads/TAO AI画册.pdf"
    doc = pymupdf.open(ORG_PDF_PATH)

    for page in doc:
        page_content = page.get_text(option="dict")

        # print("get page content")
        content_list = []
        for block in page_content["blocks"]:
            tmp_block_content_list = []
            # tmp_block_content = ""
            page.draw_rect(pymupdf.Rect(block["bbox"]), color=(1, 0, 0))

            if "lines" in block:
                for line in block["lines"]:
                    # page.draw_rect(pymupdf.Rect(line["bbox"]))
                    for span in line["spans"]:
                        # page.draw_rect(pymupdf.Rect(span["bbox"]))
                        # tmp_block_content += span["text"]
                        tmp_block_content_list.append(span["text"])
                        # pass
                tmp_block_content = " ".join(tmp_block_content_list)
                print(f"tmp_block_content:{tmp_block_content}")
                # while "  " in tmp_block_content:
                #     tmp_block_content.replace("  ", " ")
                while True:
                    if "  " in tmp_block_content:
                        tmp_block_content = tmp_block_content.replace("  ", " ")
                    else:
                        break
            content_list.append(tmp_block_content)

        # tables = page.find_tables()
        # print("get tables")

        # # for table in tables.tables:
        # #     for cell in table.cells:
        # #         table.page.draw_rect(cell, color=(1, 0, 0))

        # for table in tables.tables:
        #     table.page.draw_rect(table.bbox, color=(0, 1, 0))
    doc.save("block.pdf")
