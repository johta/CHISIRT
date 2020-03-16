import openpyxl
from PIL import Image,ImageDraw,ImageFont
import textwrap
import os


font = ImageFont.truetype('/Library/Fonts/ヒラギノ角ゴシック W0.ttc', 30)

def drawCard():
    #エクセル読み込み（ファイル名・シート名はハードコード）
    workbook = openpyxl.load_workbook('cardsheet.xlsx')
    sheet = workbook["Sheet1"]

    for i in range(2,20):
        card_num = "card_"+str(i)
        print(card_num)
        id = sheet.cell(row=i, column=1).value
        title = sheet.cell(row=i, column=2).value
        cost =sheet.cell(row=i, column=5).value
        description = sheet.cell(row=i, column=3).value
        # flavor =sheet.cell(row=i, column=4).value

        im = Image.new("RGB",(708,1033),(255,255,255))
        draw = ImageDraw.Draw(im)

        draw.rectangle((2, 2, 706, 1030), fill=(255, 255, 255) ,outline=(0,0,0)) #Frame
        draw.rectangle((10, 10, 100, 90), fill=(255, 255, 255) ,outline=(0,0,0)) #Number
        draw.rectangle((110, 10, 610, 90), fill=(255, 255, 255) ,outline=(0,0,0)) #Title
        draw.rectangle((620, 10, 698, 90), fill=(255, 255, 255) ,outline=(0,0,0)) #Cost
        draw.rectangle((40, 110, 668, 500), fill=(255, 255, 255) ,outline=(0,0,0)) #illust
        draw.rectangle((40, 520, 668, 900), fill=(255, 255, 255) ,outline=(0,0,0))#description
        draw.rectangle((40, 920, 668, 1000), fill=(255, 255, 255) ,outline=(0,0,0))#flavor text

        draw.multiline_text((30, 30), card_number, fill=(0, 0, 0), font=font)

        #カード名の印字（折返し有り）
        wrap_list = textwrap.wrap(card_title, 16)
        line_counter = 0  # 行数のカウンター
        for line in wrap_list:  # wrap_listから1行づつ取り出しlineに代入
            y = line_counter*40+10  # y座標をline_counterに応じて下げる
            draw.multiline_text((120, y+5),line, fill=(0,0,0), font=font)  # 1行分の文字列を画像に描画
            line_counter = line_counter +1  # 行数のカウンターに1

        #カード効果の印字（折返し有り）
        wrap_list = textwrap.wrap(card_description, 20)
        line_counter = 0  # 行数のカウンター
        for line in wrap_list:  # wrap_listから1行づつ取り出しlineに代入
            y = line_counter*70+80  # y座標をline_counterに応じて下げる
            draw.multiline_text((50, y+550),line, fill=(0,0,0), font=font)  # 1行分の文字列を画像に描画
            line_counter = line_counter +1  # 行数のカウンターに1

        draw.multiline_text((640, 30), str(card_cost), fill=(0, 0, 0), font=font)

        #確認用
        # im.show()

        #でかすぎるのでリサイズ
        im_resize = im.resize((int(im.width*0.48), int(im.height*0.48)),Image.LANCZOS)
        im_resize.save('card/'+card_num+'.jpg', quality=95)

if __name__ == '__main__':
    drawCard()
