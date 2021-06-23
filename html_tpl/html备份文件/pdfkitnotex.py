# coding=utf-8

import sys
import io
import os
import fitz
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# pdfmetrics.registerFont(TTFont('微软雅黑', 'msyh.ttc'))heiti.TTF
pdfmetrics.registerFont(TTFont('heiti', 'simhei.ttf'))


def parse(file_path):
    doc = fitz.open(file_path)
    pageCount = doc.pageCount
    mltemp = ['01基本信息', '02检测结果', '03结果列表', '04疑似人体共生微生物列表', '05补充报告-耐药基因', '06测序数据质控', '07微生物解释说明', '08检测方法介绍', '09参考文献']
    addpagenmu = []
    for page in range(2, pageCount):
        page1 = doc.loadPage(page)
        page1text = page1.getText("text").replace(' ','')
        for d in mltemp:
            if d in page1text:
                addpagenmu.append(page-1)
                # print(page1text)
                continue
    return list(set(addpagenmu))


if __name__ == '__main__':
    filename = '0320.nj2h'
    dir_path = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(dir_path, filename+'.pdf')
    # 输出目录
    outurl = os.path.join(dir_path, filename+'-out.pdf')
    pagenmu = parse(file_path)

    print(pagenmu)

    pdf = PdfFileReader(file_path)
    pdf_writer = PdfFileWriter()
    pagesnum = pdf.getNumPages()
    # 复制首页和目录不参与增加页码
    pdf_page = pdf.getPage(0)
    pdf_writer.addPage(pdf_page)
    # 处理目录
    # can.setFillColor('#999999')
    if len(pagenmu) == 9:
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont('heiti', 13)
        mmup = 0
        for d in pagenmu:
            # can.drawString((174)*mm, (170-mmup)*mm, "- " + str(d) + " -")
            can.drawString((174)*mm, (167.5-mmup)*mm, "- " + str(d) + " -")
            # can.drawImage(self, image, x, y, width=None,height=None,mask=None)
            mmup = mmup+12
        can.save()
        packet.seek(0)

        watermark = PdfFileReader(packet)
        watermark_page = watermark.getPage(0)

        pdf_page = pdf.getPage(1)
        pdf_page.mergePage(watermark_page)
        pdf_writer.addPage(pdf_page)
    else:
        pdf_page = pdf.getPage(1)
        pdf_writer.addPage(pdf_page)
        print('目录处理失败')
    # 首页目录之后增加页码
    for page in range(2, pdf.getNumPages()):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont('heiti', 9.5)
        # can.setFillColor('#999999')
        # can.drawString((210//2)*mm, (4)*mm, "第" + str(page) + "页/共" + str(pagesnum))
        # can.drawString((210//2)*mm, (4)*mm, str(page-1) + "/" + str(pagesnum-2))
        can.drawString((194)*mm, (18.5)*mm, "- " + str(page-1) + " -")
        can.save()
        packet.seek(0)

        watermark = PdfFileReader(packet)
        watermark_page = watermark.getPage(0)

        pdf_page = pdf.getPage(page)
        pdf_page.mergePage(watermark_page)
        pdf_writer.addPage(pdf_page)

    with open(outurl, 'wb') as fh:
        pdf_writer.write(fh)