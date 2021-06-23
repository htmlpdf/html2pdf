# 导入库
import pdfkit
import os
'''将html文件生成pdf文件'''


def html_to_pdf(html, to_file):
    # 将wkhtmltopdf.exe程序绝对路径传入config对象
    path_wkthmltopdf = r'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    options = {
        'page-size': 'A4',
        'margin-top': '25.4mm',
        'margin-right': '0mm',
        'margin-bottom': '25.4mm',
        'margin-left': '0mm',
            # 'orientation':'Landscape',#横向
        'encoding': "UTF-8",
        # 'no-outline': None,
        'zoom': '1',
            # 'footer-right':'[page]'
        '--enable-local-file-access': '--enable-local-file-access',
        # '--header-right': '页眉右侧文字',
        # '--header-line': '--header-line', # 页眉下划线
        # '--header-spacing': 5, # 页眉距离内容高度
    }
    # 生成pdf文件，to_file为文件路径
    pdfkit.from_file(html, to_file, options=options, configuration=config)
    print('完成')

html_to_pdf('nmgs1.html', 'out_2.pdf')

'''将网页生成pdf文件'''


# def url_to_pdf(url, to_file):
#     # 将wkhtmltopdf.exe程序绝对路径传入config对象
#     path_wkthmltopdf = r'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
#     config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
#     # 生成pdf文件，to_file为文件路径
#     options = {
#         'page-size': 'A4',
#         'margin-top': '0mm',
#         'margin-right': '0mm',
#         'margin-bottom': '0mm',
#         'margin-left': '0mm',
#         # 'orientation':'Landscape',#横向
#         'encoding': "UTF-8",
#         'no-outline': None,
#         'zoom': '1.1',
#         # 'footer-right':'[page]'
#         '--enable-local-file-access': '--enable-local-file-access',
#         # '--header-right': '页眉右侧文字',
#         # '--header-line': '--header-line', # 页眉下划线
#         # '--header-spacing': 5, # 页眉距离内容高度
#         # 'header-left'  : 'something',
#         # 'header-right' : '[section]',
#         # 'footer-right' : '[page]',
#     }
#     pdfkit.from_url(url, to_file, options=options, configuration=config,)
#     print('完成')


# url_to_pdf(r'http://localhost:8080/#/nmgs', 'out_11.pdf')
