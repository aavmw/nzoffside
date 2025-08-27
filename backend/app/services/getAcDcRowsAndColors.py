from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import time
import json
from ..config import get_settings

settings = get_settings()

with open(settings.GOOGLE_CREDS_PATH) as f:
        credentials = json.loads(f.read())

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def appendTable(id, sheetname, values, value_input_option):
    creds = Credentials.from_service_account_info(credentials, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    body = {"values": values}
    service.spreadsheets().values().append(
        spreadsheetId=id,
        range=sheetname,
        valueInputOption=value_input_option,
        body=body,
    ).execute()


def formatColorsToHex(formatting):
    rows_color = []
    for row in range(len(formatting)):
        cell_colors = []
        for cell in range(len(formatting[row])):
            red = int(formatting[row][cell].get("red", 0) * 255)
            green = int(formatting[row][cell].get("green", 0) * 255)
            blue = int(formatting[row][cell].get("blue", 0) * 255)
            cell_colors.append(f"#{red:02x}{green:02x}{blue:02x}")
        rows_color.append(cell_colors)
    return rows_color


def clearSheet(id, sheetname):
    creds = Credentials.from_service_account_info(credentials, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    service.spreadsheets().values().clear(
        spreadsheetId=id,
        range=sheetname,
    ).execute()


def getValuesAndFormatting(sheetId, sheetName, tablename):
    creds = Credentials.from_service_account_info(credentials, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sheet = (
        service.spreadsheets()
        .get(spreadsheetId=sheetId, ranges=[sheetName], includeGridData=True)
        .execute()
    )

    row_data = sheet["sheets"][0]["data"][0]["rowData"]

    data_values, formatting, fontsRGB = [], [], []

    for row in row_data:
        try:
            if "AC-DC" in row["values"][10].get("formattedValue", ""):
                row_values, row_formatting, font_formatting = [tablename], [], []
                for cell in row.get("values", []):
                    row_values.append(cell.get("formattedValue", ""))
                    format_data = cell.get("effectiveFormat", {})
                    background_color = format_data.get("backgroundColor", {})
                    row_formatting.append(background_color)
                    text_format = format_data.get("textFormat", {})
                    font_color = text_format.get("foregroundColor", {})
                    font_formatting.append(font_color)
                data_values.append(row_values)
                formatting.append(row_formatting)
                fontsRGB.append(font_formatting)
        except KeyError:
            continue
    formatting = formatColorsToHex(formatting)
    formatting = [[i[0]] + i for i in formatting]
    fontsRGB = formatColorsToHex(fontsRGB)
    fontsRGB = [[i[0]] + i for i in fontsRGB]
    return data_values, formatting, fontsRGB


def getFull():

    spreadsheetsMap = {
        "summary_ac_dc": "1ZP7seBw0PuEluUB8nli8qQrdRODGDrImNaCiNZrQkXM",
        "K Orders": "1gHXr2kNVZkdeLCzXGTMLgUKYdXKy5p1a1f7OVRqaJaE",
        "A Orders": "1rQFIQLTpOQUtylIQEB3ky1firEnVGgWQ47CgEuU67ho",
        "B Orders": "15GIeyf9H3IbLOQz7JRFf58yAm14lsa1lLnrWI0ZmYYE",
        "L Orders": "1DN6r1iMWXmIgoTfQcSCeRO6WO97GYTXrKwez_2DrVyU",
        "Y Orders": "1lvgTVF-9HT1CmvLb482XNzRcitkpY-FLxVP-wGRRcr4",
        "Z Orders": "14YxNRFIWjF5YeR088zIH2u2o0XcibWsAC45x9Zcwmlc",
        "Orders (AC DC)": "1prkJPH9sGXwkavYBU3dPx3uOGcocDSmI7oYEj3GI3W8",
        "Q Orders": "1FT3kM597Pm2RgxuQTgmh_uQT4hmlYk8FPLLr_fIbLcQ",
        "Y2 Orders": "1ISaD-LQo5rOwWhG1D-Ty-0ZltgMeBRYf5oIOQ9s8DJ8",
        "R Orders": "1k31s2zXwXNGHT8i_-CNYcOTu-ibcMDOoBThHSnCXxbY",
    }

    listsMap = [
        ("K Orders", "ORDERS!A:AA"),
        ("K Orders", "ORDERS AC!A:AA"),
        ("A Orders", "ORDERS!A:AA"),
        ("B Orders", "ORDERS!A:AA"),
        ("Z Orders", "ORDERS AC!A:AA"),
        ("Z Orders", "TOOL AC!A:AA"),
        ("L Orders", "ORDERS!A:AA"),
        ("L Orders", "ORDERS AC!A:AA"),
        ("L Orders", "TOOL AC!A:AA"),
        ("Y Orders", "ORDERS AC!A:AA"),
        ("Y Orders", "TOOL AC!A:AA"),
        ("Orders (AC DC)", "for_summary!A:AA"),
        ("Q Orders", "ORDERS AC!A:AA"),
        ("Y2 Orders", "ORDERS AC!A:AA"),
        ("Y2 Orders", "TOOL AC!A:AA"),
        ("R Orders", "TOOL AC!A:AA"),
    ]

    clearSheet(spreadsheetsMap["summary_ac_dc"], "values")
    clearSheet(spreadsheetsMap["summary_ac_dc"], "background_colors")
    clearSheet(spreadsheetsMap["summary_ac_dc"], "font_colors")
    data, colors, fonts = [], [], []
    for i in listsMap:
        time.sleep(1)
        try:
            source = getValuesAndFormatting(spreadsheetsMap[i[0]], i[1], i[0])
            data.extend(source[0])
            colors.extend(source[1])
            fonts.extend(source[2])
        except Exception as err:
            pass
            print(f"{i[0]} {i[1]} failed with {err}")

    appendTable(spreadsheetsMap["summary_ac_dc"], "values", data, "RAW")
    appendTable(spreadsheetsMap["summary_ac_dc"], "background_colors", colors, "RAW")
    appendTable(spreadsheetsMap["summary_ac_dc"], "font_colors", fonts, "RAW")
