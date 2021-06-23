#! coding=utf-8
import os,sys
import io
import json
import fitz
import pdfkit
from docxtpl import DocxTemplate, RichText
from jinja2 import PackageLoader, Environment, FileSystemLoader
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
# 设置目录及页码字体msyh.ttc
pdfmetrics.registerFont(TTFont('heiti', '/data/softwares/mngs_scripts/report/html2pdf/html_tpl/fonts/simhei.ttf'))
pdfmetrics.registerFont(TTFont('微软雅黑', '/data/softwares/mngs_scripts/report/html2pdf/html_tpl/fonts/msyh.ttc'))
pdfmetrics.registerFont(TTFont('宋体', '/data/softwares/mngs_scripts/report/html2pdf/html_tpl/fonts/simsun.ttc'))

texdata = {}
with open(sys.argv[1], 'r',encoding='utf-8') as infile:
    texdata = json.load(infile)

# html转pdf格式配置 
options = {
        'page-size': 'Letter',
        'margin-top': '25.4mm',
        'margin-right': '19.1mm',
        'margin-bottom': '25.4mm',
        'margin-left': '19.1mm',
        'encoding': "UTF-8",
        'no-outline': None,
        'zoom': '1.1',
        '--enable-local-file-access': '--enable-local-file-access',
        '--javascript-delay': 10,
        # '--load-error-handling': 'ignore',
        # '--load-media-error-handling': 'ignore',
        
    }
# 封面与目录页眉页脚
options1 = {
    # 'page-size': 'A4',
    'page-size': 'Letter',
    'margin-top': '0mm',
    'margin-right': '0mm',
    'margin-bottom': '0mm',
    'margin-left': '0mm',
    'encoding': "UTF-8",
    'zoom': '1.2',
    '--enable-local-file-access': '--enable-local-file-access',
    # '--load-error-handling': 'ignore',
    # '--load-media-error-handling': 'ignore',
    '--javascript-delay': 10,
}
# 获取当前绝对地址
dir_path = os.path.dirname(os.path.abspath(__file__))

tpl_dir = f'/data/softwares/mngs_scripts/report/html2pdf/html_tpl/'
outdir = sys.argv[2]

# 初始html模版填充 传入模版名称
def html_set_data(filename):
    file_path = filename+'.html'
    out_path = filename+'to.html'
    #path_dir = os.getcwd()
    loader = FileSystemLoader(searchpath=tpl_dir)
    env = Environment(loader=loader)
    template = env.get_template(file_path) # 模板文件 highSpecial
    buf = template.render(texdata)
    with open(os.path.join(outdir, out_path), "w", encoding="utf-8") as fp:
        fp.write(buf)

# 运行html模版填充 传入
def runhtml_set_data(modename,hsyname):
    #texdata['ybit'] = setwname
    # 封面及目录html模版
    filename = hsyname[1]
    html_set_data(filename)
    # 页眉页脚html模版 
    filename = hsyname[0]
    html_set_data(filename)
    if len(hsyname) == 3:
        filename = hsyname[2]
        html_set_data(filename)
    # 主内容html模版 
    filename = modename
    html_set_data(filename)


# html转pdf原始 传入绝对rul及配置options
def html_to_pdf(filename, options):
    file_path = os.path.join(outdir, filename+'.html')
    out_path = os.path.join(outdir, filename+'-out.pdf')
    # 将wkhtmltopdf.exe程序绝对路径传入config对象
    #path_wkthmltopdf = r'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
    path_wkthmltopdf = r'/usr/local/bin/wkhtmltopdf'
    config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    # 生成pdf文件，to_file为文件路径
    try:
        pdfkit.from_file(file_path, out_path, options=options, configuration=config)
        # 生成之后删除转换完的html文件减少占用资源
        os.remove(file_path)
    except Exception as e:
        if os.path.exists(file_path):
            os.remove(file_path)

# 运行三个文件html转pdf
def runhtmlhtml_to_pdf(filenametop,hsyname):
    try:
        filename = hsyname[1]+'to'
        html_to_pdf(filename, options1)
        # 页眉页脚html转pdf
        filename = hsyname[0]+'to'
        html_to_pdf(filename, options1)
        # xy最后一页特殊处理
        if len(hsyname) == 3:
            filename = hsyname[2]+'to'
            html_to_pdf(filename, options1)
        # 主要内容html转pdf
        filename = filenametop+'to'
        html_to_pdf(filename, options)
    except Exception as e:
        print('sy报错:')
        try:
            # 页眉页脚html转pdf
            filename = hsyname[0]+'to'
            html_to_pdf(filename, options1)
            # xy最后一页特殊处理
            if len(hsyname) == 3:
                filename = hsyname[0]+'to'
                html_to_pdf(filename, options1)
             # 主要内容html转pdf
            filename = filenametop+'to'
            html_to_pdf(filename, options)
        except Exception as e:
            print('报错:')
            try:
                # xy最后一页特殊处理
                if len(hsyname) == 3:
                    filename = hsyname[0]+'to'
                    html_to_pdf(filename, options1)
                # 主要内容html转pdf
                filename = filenametop+'to'
                html_to_pdf(filename, options)
            except Exception as e:
                print('报错:')
                try:
                    # 主要内容html转pdf
                    filename = filenametop+'to'
                    html_to_pdf(filename, options)
                except Exception as e:
                    print('报错:')
            

# 添加封面及目录file_path参数为封面与目录的pdf文件路径
def addfyml(file_path,file_path1,outurl):

    pdf = PdfFileReader(file_path)
    pdf1 = PdfFileReader(file_path1)
    # 创建新的pdf
    pdf_writer = PdfFileWriter()

    pagesnum = pdf.getNumPages()
    pagesnum1 = pdf1.getNumPages()
    # 0413.xy之后一页特殊处理
    if file_path1 == outurl:
        # 先复制主要内容页面
        for page in range(0, pagesnum1):
            pdf_page = pdf1.getPage(page)
            pdf_writer.addPage(pdf_page)
        # 复制封面目录页面
        for page in range(0, pagesnum):
            pdf_page = pdf.getPage(page)
            pdf_writer.addPage(pdf_page)
    else:
        # 先复制封面目录页面
        for page in range(0, pagesnum):
            pdf_page = pdf.getPage(page)
            pdf_writer.addPage(pdf_page)
        # 复制主要内容页面
        for page in range(0, pagesnum1):
            pdf_page = pdf1.getPage(page)
            pdf_writer.addPage(pdf_page)
    # 保存为文件
    with open(outurl, 'wb') as fh:
        pdf_writer.write(fh)

# 单独添加水印
def addwatermarkn(file_path,addstr):
    pdf_watermark = PdfFileReader(file_path)
    pdf_writer = PdfFileWriter()
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.translate(30 * mm, 100 * mm)
    can.setFont('heiti', 80)
    # 指定描边的颜色
    can.setStrokeColorRGB(0, 1, 0)
    # 指定填充颜色
    can.setFillColorRGB(0, 1, 0)
    # 画一个矩形
    # can.rect(cm, cm, 7*cm, 17*cm, fill=1)
    # 旋转45度,坐标系被旋转
    can.rotate(30)
    # 指定填充颜色
    can.setFillColorRGB(0, 0, 0, 0.1)
    # 设置透明度,1为不透明
    # can.setFillAlpha(0.1)
    # 画几个文本,注意坐标系旋转的影响
    can.drawString(30 * mm, 0 * mm, addstr)
    can.setFillAlpha(0.6)
    can.save()
    packet.seek(0)
    watermark = PdfFileReader(packet)
    watermark_page = watermark.getPage(0)
    pdf_page = pdf_watermark.getPage(0)
    pdf_page.mergePage(watermark_page)
    pdf_writer.addPage(pdf_page)
    with open(file_path, 'wb') as fh:
        pdf_writer.write(fh)

# 添加页眉页脚及水印
def addyeyj(file_path,file_path1,outurl,watermark):
    # 是否添加水印
    if watermark != '':
        addwatermarkn(file_path,watermark)
        pdf_watermark = PdfFileReader(file_path)
    else:
        pdf_watermark = PdfFileReader(file_path)

    pdf1 = PdfFileReader(file_path1)
    # 创建新的pdf
    pdf_writer = PdfFileWriter()
    pagesnum1 = pdf1.getNumPages()

    # 复制主要内容页面
    for page in range(0, pagesnum1):
        pdf_page = pdf1.getPage(page)
        pdf_page.mergePage(pdf_watermark.getPage(0))
        # pdf_page.compressContentStreams() # 压缩内容
        pdf_writer.addPage(pdf_page)
    # 保存为文件
    with open(outurl, 'wb') as fh:
        pdf_writer.write(fh)

# 添加封面及目录页眉页脚并合并为一个文件
def addfymlyeyj(modename,setwname,hsyname,watermark):
    ymyj_path = os.path.join(outdir, hsyname[0]+'to-out.pdf')
    fyml_path = os.path.join(outdir, hsyname[1]+'to-out.pdf')
    con_path = os.path.join(outdir, modename+'to-out.pdf')
    # 输出目录
    outurl = os.path.join(outdir, setwname+'.pdf')
    addyeyj(ymyj_path,con_path,con_path,watermark)
    addfyml(fyml_path,con_path,outurl)
    if len(hsyname) == 3:
        fyml1_path = os.path.join(outdir, hsyname[2]+'to-out.pdf')
        addfyml(fyml1_path,outurl,outurl)
        os.remove(fyml1_path)
    # 完成后删除文件以免占用资源
    os.remove(ymyj_path)
    os.remove(fyml_path)
    os.remove(con_path)

# 扫描目录bcnum扫描也补偿
def parse(file_path,mltempstr):
    doc = fitz.open(file_path)
    pageCount = doc.pageCount
    # mltempstr = mltempstr
    addpagenmu = []
    for page in range(2, pageCount):
        page1 = doc.loadPage(page)
        page1text = page1.getText("text").replace(' ','')
        for d in mltempstr:
            if d in page1text:
                addpagenmu.append(page-1)
                # print(page1text)
                continue
    return list(set(addpagenmu))

# 添加目录和页码
def parseaddmlym(filename,setout_name):
    file_path = os.path.join(outdir, filename+'.pdf')
    # 输出目录
    outurl = os.path.join(outdir, setout_name+'.pdf')
    pdf = PdfFileReader(file_path)
    pdf_writer = PdfFileWriter()
    pagesnum = pdf.getNumPages()
    # 复制首页和目录不参与增加页码
    pdf_page = pdf.getPage(0)
    pdf_writer.addPage(pdf_page)
    # 处理目录 0201.boao不处理目录
    if filename == '0201.boao':
        # 首页目录之后增加页码
        for page in range(1, pdf.getNumPages()):
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            can.setFont('微软雅黑', 7)
            can.setFillColor('#595959', alpha=None)
            can.drawString((191)*mm, (21)*mm, "第 " + str(page) + " 页")
            can.save()
            packet.seek(0)

            watermark = PdfFileReader(packet)
            watermark_page = watermark.getPage(0)

            pdf_page = pdf.getPage(page)
            pdf_page.mergePage(watermark_page)
            pdf_writer.addPage(pdf_page)
    else:
        # 扫描目录
        pagenmu = parse(file_path,mltempobj[filename])
        # 每个目录是否都检索到
        print(pagenmu)
        if len(pagenmu) == len(mltempobj[filename]):
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            can.setFont('heiti', 13)
            # 为改模版目录添加头坐标补偿参数
            if filename == '0201.mz' :
                topcompensate = 69.3
            elif filename == '0525.xy' or filename == '0618.cd':
                topcompensate = 69.3
            else:
                topcompensate = 0
            mmup = 0
            for d in pagenmu:
                can.drawString((184)*mm, ((142.7+topcompensate)-mmup)*mm, "- " + str(d) + " -")
                mmup = mmup+12.1
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
            if filename == '0525.xy' or filename == '0618.cd':
                can.setFont('heiti', 7.5)
                can.setFillColor('#595959', alpha=None)
                if (page != pdf.getNumPages()-1):
                    can.drawString((190)*mm, (21)*mm, "第 " + str(page) + " 页")
            else:
                can.setFont('heiti', 9.5)
                can.drawString((200)*mm, (20.8)*mm, "- " + str(page-1) + " -")
            can.save()
            packet.seek(0)

            watermark = PdfFileReader(packet)
            watermark_page = watermark.getPage(0)

            pdf_page = pdf.getPage(page)
            pdf_page.mergePage(watermark_page)
            pdf_writer.addPage(pdf_page)
    with open(outurl, 'wb') as fh:
        pdf_writer.write(fh)
    # 完成后删除文件以免占用资源
    os.remove(file_path)

# 复制pdf到某目录并添加页码
def copypdfpath(file_path,out_path):

    pdf = PdfFileReader(file_path)
    pdf_writer = PdfFileWriter()
    pagesnum = pdf.getNumPages()
    for page in range(0, pagesnum):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont('宋体', 9)
        can.drawString((107)*mm, (18)*mm,  str(page+1))
        can.save()
        packet.seek(0)
        watermark = PdfFileReader(packet)
        watermark_page = watermark.getPage(0)
        pdf_page = pdf.getPage(page)
        pdf_page.mergePage(watermark_page)
        pdf_writer.addPage(pdf_page)
    # 保存为文件
    with open(out_path, 'wb') as fh:
        pdf_writer.write(fh)

# 目录结构
mltemp = ['01基本信息', '02检测结果', '03结果列表', '04疑似人体共生微生物列表', '05补充报告-耐药基因', '06测序数据质控', '07微生物解释说明', '08检测方法介绍', '09参考文献']
mltemp1 = ['01基本信息', '02检测结果', '03结果列表', '04补充报告-微生物列表', '05补充报告-耐药基因',  '06微生物解释说明', '07检测方法介绍', '08参考文献']
mltemp2 = ['一、基本信息', '二、检测结果', '三、结果列表', '四、补充报告-微生物列表', '五、补充报告-耐药基因',  '六、微生物解释说明', '七、检测方法介绍', '八、参考文献']
mltemp3 = ['一、基本信息', '二、检测结果', '三、结果列表', '四、疑似人体共生微生物列表', '五、补充报告-耐药基因',  '六、测序数据质控', '七、微生物解释说明', '八、检测方法介绍', '九、参考文献']
mltempobj = {
    '0413.xy': mltemp, '0325.aja': mltemp, '0325.zju': mltemp, '0320.nj2h': mltemp1, '0201.fzch': mltemp1, '0201.hy': mltemp1, '0201.mz': mltemp2,
    '0525.xy': mltemp3, '0618.ql': mltemp, '0618.cd': mltemp3,
}
# 不同版本对于的使用模版对照表
contrastobj = {
    '0413.xy': '0325_0413_stencil', '0325.aja': '0325_0413_stencil', '0325.zju': '0325_0413_stencil', '0320.nj2h': '0320_nj2h_stencil',
    '0201.fzch': '0320_nj2h_stencil','0201.hy': '0320_nj2h_stencil', '0201.mz': '0201_mz_stencil','0201.boao': '0201_boao_stencil',
    '0420.ry': '0420_ry_stencil', '0525.xy': '0525_xy_stencil', '0618.ql': '0325_0413_stencil', '0618.cd': '0525_xy_stencil',
}
# 对于模版的对于封面及页眉页脚模版对照表
fmluymyjobj = {
    '0413.xy': ['hderfter','symltop','hderfter_last'], '0325.aja': ['hderfter','symltop'], '0325.zju': ['hderfter','symltop'], '0320.nj2h': ['hderfter','symltop'],
    '0201.fzch': ['hderfter','symltop'],'0201.hy': ['hderfter','symltop'], '0201.mz': ['hderfter1','symltop1'], '0201.boao': ['hderfter2','symltop2'],
    '0525.xy': ['hderfter3','symltop3','hderfter_last'],'0618.ql':['hderfter4','symltop4','hderfter_last1'],'0618.cd': ['hderfter3','symltop3','hderfter_last']
}

if __name__ == '__main__':
    # 选择要出的么办文件名称： 0413.xy/0325.aja/0325.zju/0320.nj2h/0201.fzch/0201.hy/0201.mz/0201.boao/0420.ry/0525.xy
    setwname = texdata["ybit"]
    # 水印文字变量
    watermark = texdata["shuiyin"]
    # 自定义文件名变量
    setout_name = f'{texdata["report_id"]}_{texdata["department_id"]}_{texdata["name"]}_mNGS检测报告'
    modename = contrastobj[setwname]
    if setwname == '0420.ry':
        html_set_data(modename)
        html_to_pdf(modename+'to', options)
        copypath = os.path.join(outdir, modename+'to-out.pdf')
        outpath = os.path.join(outdir, setout_name+'.pdf')
        copypdfpath(copypath,outpath)
        # 完成后删除文件以免占用资源
        os.remove(copypath)
    else:
        # 三个html填充数据
        runhtml_set_data(modename,fmluymyjobj[setwname])
        runhtmlhtml_to_pdf(modename,fmluymyjobj[setwname])
        addfymlyeyj(modename,setwname,fmluymyjobj[setwname],watermark)
        parseaddmlym(setwname,setout_name)
