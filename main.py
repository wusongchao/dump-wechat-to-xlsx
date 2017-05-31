# -*- coding:utf-8 -*-
import xlsxwriter
import csv


EXPORTED_DATA_PATH = "ExportedData/"
XLSX_OUT_PATH = "out/"

MESSAGE_TYPE = {
    "34": "voice",
    "3": "image",
    "1": "text"
}

IMG_INFO_INFO = {
    "msgSvrId": "",
    "bigImgPath": "",
    "thumbImgPath": "",
    "createtime": ""
}

IMG_INFO_LIST = []

RCONTACT_INFO = {
    "username": "",
    "nickname": ""
}

RCONTACT_LIST = []

MESSAGE_INFO = {
    "msgId": "",
    "msgSvrId": "",
    "type": "",
    "status": "",
    "createTime": "",
    "talker": "",
    "content": "",
    "imgPath": "",
    "talkerId": "",
}

MESSAGE_LIST = []


def dump_img_info(pathname):
    with open(pathname) as img_info2_f:
        reader = csv.DictReader(img_info2_f)

        for row in reader:
            for_append = {
                "msgSvrId": row["msgSvrId"],
                "bigImgPath": row["bigImgPath"],
                "thumbImgPath": row["thumbImgPath"],
                "createtime": row["createtime"]
            }
            IMG_INFO_LIST.append(for_append)


def dump_rcontact(pathname):
    RCONTACT_LIST.append({
        "username": "placeholder",
        "nickname": "placeholder"
    })

    with open(pathname, mode="rt", encoding="ISO-8859-1") as rcontact_f:
        reader = csv.DictReader(rcontact_f)

        for row in reader:
            for_username_append = row["username"]
            for_nickname_append = row["nickname"]

            try:
                for_nickname_append = bytes(for_nickname_append, "ISO-8859-1").decode("gb2312")
            except UnicodeDecodeError as o:
                pass

            RCONTACT_LIST.append({
                "username": for_username_append,
                "nickname": for_nickname_append
            })


def dump_message(pathname):
    with open(pathname, mode="rt", encoding="ISO-8859-1") as message_f:
        reader = csv.DictReader(message_f)

        for row in reader:
            for_content_append = row["content"]

            try:
                for_content_append = bytes(for_content_append, "ISO-8859-1").decode("gb2312")
            except UnicodeDecodeError as o:
                pass

            for_append = {}

            for key in MESSAGE_INFO.keys():
                if key == "content":
                    for_append[key] = for_content_append
                else:
                    for_append[key] = row[key]

            MESSAGE_LIST.append(for_append)


def search_img_by_thumbnail(thumbnail_name):
    for item in IMG_INFO_LIST:
        if item["thumbImgPath"] == thumbnail_name:
            return item["bigImgPath"]

    return ""


def generate_xlsx(pathname):
    wb = xlsxwriter.Workbook(pathname)
    ws = wb.add_worksheet()
    # number_format = wb.add_format({"num_format": ""})

    header = ["msgId", "msgSvrId", "talkerUsername", "talkerNickname",
              "type", "content", "sourcePath", "createTime"]
    ws.write_row("A1", header)

    row = 1
    col = 0

    for line in MESSAGE_LIST:
        ws.write(row, col, line["msgId"])
        ws.write(row, col+1, line["msgSvrId"])
        ws.write(row, col+2, line["talker"])
        ws.write(row, col+3, RCONTACT_LIST[int(line["talkerId"])]["nickname"])

        if line["type"] in MESSAGE_TYPE.keys():
            ws.write(row, col+4, MESSAGE_TYPE[line["type"]])
        else:
            ws.write(row, col+4, "")
        ws.write(row, col+5, line["content"])

        if line["type"] == "3":  # represent img
            ws.write(row, col+6, search_img_by_thumbnail(line["imgPath"]))

        ws.write(row, col+7, line["createTime"])

        row += 1

    wb.close()


def main():
    dump_img_info(EXPORTED_DATA_PATH + "ImgInfo2.csv")
    dump_rcontact(EXPORTED_DATA_PATH + "rcontact.csv")
    dump_message(EXPORTED_DATA_PATH + "message.csv")
    generate_xlsx(XLSX_OUT_PATH + "out.xlsx")


if __name__ == '__main__':
    main()
