import os
import json
from docxtpl import DocxTemplate, RichText
from jinja2 import PackageLoader, Environment, FileSystemLoader

highBacteria = {
    'name': "细菌",'species': [
        {'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'},
        {'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'},
    ],
}
highVirus = {
    'name': "病毒",'species': [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}
highFungi = {
    'name': "真菌",'species': [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}
highParasite = {
    'name': "寄生虫",'species': [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}
highSpecial = {
    'name': "特殊病原体（包括分枝杆菌、支原体/衣原体等）",'species': [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}

lowBacteria  = {
    'name': "细菌",'species': [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}
lowVirus = {
    'name': "病毒",'species': [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}
lowFungi = {
    'name': "真菌",'species': [
        {'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'},
        {'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'},
        {'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}
    ],
}
lowParasite = {
    'name': "寄生虫",'species': [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}
lowSpecial = {
    'name': "特殊病原体（包括分枝杆菌、支原体/衣原体等）",'species':  [{'species_cn':'耶氏肺孢子菌','species_e':'Pneumocystis jirovecii'}],
}
hightyoe = [highBacteria,highVirus,highFungi,highParasite,highSpecial]
lowtyoe = [lowBacteria,lowVirus,lowFungi,lowParasite,lowSpecial]
# 模拟数据

# 结果列表自有数据
results_list = ['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）']
# 结果列表填充数据
bacteriaList = [
    {'genus':'克雷伯氏菌属','genus_e':'Klebsiella', 'gcount':'9', 'area':[{'type':'G-', 'species':'产气克雷伯氏菌','species_e':'Klebsiella aerogenes', 'scount':'7', 'abundance':'0.01%', 'focus':'低'}]},
    {'genus':'克雷伯氏菌属','genus_e':'Klebsiella', 'gcount':'9', 'area':[{'type':'G-', 'species':'产气克雷伯氏菌','species_e':'Klebsiella aerogenes', 'scount':'7', 'abundance':'0.01%', 'focus':'低'}]},
]
virusList = [{'genus':'淋巴滤泡病毒属','genus_e':'Lymphocryptovirus', 'gcount':'5,543', 'area':[{'type':'DNA', 'species':'EBV','species_e':'Human gammaherpesvirus 4', 'scount':'5,517', 'abundance':'3.979%', 'focus':'高'}]}]
fungiList = [
{'genus':'Klebsiella','genus_e':'Klebsiella', 'gcount':'-', 'area':[{'species':'Klebsiella','species_e':'Klebsiella', 'scount':'7', 'abundance':'-', 'focus':'-'}]},
{'genus':'Klebsiella','genus_e':'Klebsiella', 'gcount':'-', 'area':[{'species':'Klebsiella','species_e':'Klebsiella', 'scount':'7', 'abundance':'-', 'focus':'-'}]}
]
parasiteList = [{'genus':'Klebsiella','genus_e':'Klebsiella', 'gcount':'-', 'area': [{'type':'-', 'species':'Klebsiella','species_e':'Klebsiella', 'scount':'7', 'abundance':'-', 'focus':'-'}]}]
specialList = [{'genus':'Klebsiella','genus_e':'Klebsiella','gcount':'-', 'area':[{'type':'-', 'species':'Klebsiella','species_e':'Klebsiella', 'scount':'7', 'abundance':'-', 'focus':'-'}]}]
listData = [bacteriaList, virusList, fungiList, parasiteList, specialList]
# 微生物列表数据
areanet = [
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '细菌', 'microbe': '牙髓卟啉单胞菌','microbe_e':'Porphyromonas endodontalis', 'count': 4590, 'note': '疑似背景微生物'},
    {'type': '病毒', 'microbe': '人疱疹病毒7型','microbe_e':'Human betaherpesvirus 7', 'count': 1, 'note': '疑似背景微生物'}
]
# 微生物解释说明数据
descriptions = [
    {'url_name':'img/ZZ210811.Corynebacterium_striatum.png','name': '牙髓卟啉单胞菌', 'english': 'Porphyromonas endodontalis', 'explain': '牙髓卟啉单胞菌是与牙髓感染和牙源性脓肿有特殊关系的不解糖菌株。牙髓卟啉单胞菌是感染根管的优势菌,在牙髓、根尖周组织病的发生、发展中起着重要的作用,并与牙髓坏死、根尖部疼痛肿胀、叩痛等症状和窦道形成密切相关。'},
    {'url_name':'img/ZZ210811.Corynebacterium_striatum.png','name': '牙髓卟啉单胞菌', 'english': 'Porphyromonas endodontalis', 'explain': '牙髓卟啉单胞菌是与牙髓感染和牙源性脓肿有特殊关系的不解糖菌株。牙髓卟啉单胞菌是感染根管的优势菌,在牙髓、根尖周组织病的发生、发展中起着重要的作用,并与牙髓坏死、根尖部疼痛肿胀、叩痛等症状和窦道形成密切相关。'},
    {'url_name':'img/ZZ210811.Corynebacterium_striatum.png','name': '牙髓卟啉单胞菌', 'english': 'Porphyromonas endodontalis', 'explain': '牙髓卟啉单胞菌是与牙髓感染和牙源性脓肿有特殊关系的不解糖菌株。牙髓卟啉单胞菌是感染根管的优势菌,在牙髓、根尖周组织病的发生、发展中起着重要的作用,并与牙髓坏死、根尖部疼痛肿胀、叩痛等症状和窦道形成密切相关。'},
    {'url_name':'-','name': '牙髓卟啉单胞菌', 'english': 'Porphyromonas endodontalis', 'explain': '牙髓卟啉单胞菌是与牙髓感染和牙源性脓肿有特殊关系的不解糖菌株。牙髓卟啉单胞菌是感染根管的优势菌,在牙髓、根尖周组织病的发生、发展中起着重要的作用,并与牙髓坏死、根尖部疼痛肿胀、叩痛等症状和窦道形成密切相关。'},
]
# 微生物参考说明数据 
papers = [
    {'namecon':'汤亚玲. 牙髓卟啉单胞菌的生物学特性与致病性[J]. 国外医学.口腔医学分册, 2003.','conbit':True},
    {'namecon':'Downes J, Wade W G. Peptostreptococcus stomatis sp. nov., isolated from the human oral cavity[J]. International journal of systematic and evolutionary microbiology, 2006, 56(4): 751-754.','conbit':False},
    {'namecon':'关素敏. 中间普氏菌增殖机制及对慢性牙周炎致病作用的研究[D]. 第四军医大学, 2008.','conbit':True},
    {'namecon':'Bahrani-Mougeot F K, Paster B J, Coleman S, et al. Molecular analysis of oral and respiratory bacterial species associated with ventilator-associated pneumonia[J]. J Clin Microbiol, 2007, 45(5): 1588-1593.','conbit':False},
    {'namecon':'Henderson A , Wall D . Streptococcus milleri liver abscess presenting as fulminant pneumonia.[J]. Australian & New Zealand Journal of Surgery, 2010, 63(3):237-240.','conbit':False},
    {'namecon':'周村，林伟. 齿垢密螺旋体. 国外医学口腔医学分册, 2000, 27:4.','conbit':True},
    {'namecon':'金艳, 张春和, 陈东科,等. 198株黏液罗氏菌的临床分离情况及耐药性分析[J]. 检验医学, 2008(5):494-496.','conbit':True},
    {'namecon':'Nakazawa F, Poco SE, Sato M, Ikeda T, Kalfas S, Sundqvist G, Hoshino E. Taxonomic characterization of Mogibacterium diversum sp. nov. and Mogibacterium neglectum sp. nov., isolated from human oral cavities. Int J Syst Evol Microbiol. 2002 ,52(Pt 1):115-122.','conbit':False},
    {'namecon':'OKUDA K, KATO T, SHIOZU J, et al. Bacteroides hepurinolyticus sp. nov. Isolated from Humans with Periodontitis. INTERNATIONAL JOURNAL OF SYSTEMATIC BACTERIOLO, 1985, 35(4): 438-442.','conbit':False},
    {'namecon':'林玉玲, 林建芬, 陈峰. 呼吸道标本中1株奇异劳特普罗菌的分离及鉴定。临床检验杂志, 2018,36(10):721-724.','conbit':True},
    {'namecon':'丘小慧, 王进, 黄燕宁, 等. 疑似具核梭杆菌和直肠弯曲菌致脑膜炎1例. 中国感染与化疗杂志, 2019, 19(3): 315-318.','conbit':True},
    {'namecon':'喻国灿，徐旭东，叶波. 麻疹孪生球菌致肺空洞和肺脓肿一例. 中华传染病杂志，2017,35(4):245-246.','conbit':True},
    {'namecon':'Gorospe L, Bermudez-Coronel-Prats I, Gomez-Barbosa C F, et al. Parvimonas micra chest wall abscess following transthoracic lung needle biopsy[J]. The Korean journal of internal medicine, 2014, 29(6): 834-837.','conbit':False},
    {'namecon':'Himeji D , Hara S , Kawaguchi T , et al. A Case of Pulmonary Actinomyces graevenitzii Infection Diagnosed by Bronchoscopy using Endobronchial Ultrasonography with a Guide Sheath[J]. Internal Medicine, 2018.','conbit':False},
    {'namecon':'Goker, M. , Held, B. , Lucas, S. , Nolan, M. , Yasawong, M. , & Tijana, G. D. R. , et al. (2010). Complete genome sequence of olsenella uli type strain (vpi d76d-27ct). Standards in Genomic Sciences, 3(1), 76.','conbit':False},
    {'namecon':'Trofa D, Gácser A, Nosanchuk J D. Candida parapsilosis, an emerging fungal pathogen[J]. Clinical microbiology reviews, 2008, 21(4): 606-625.','conbit':False},
    {'namecon':'Palser A L, Grayson N E, White R E, et al. Genome diversity of Epstein-Barr virus from multiple tumor types and normal infection[J]. Journal of virology, 2015, 89(10): 5222-5237.','conbit':False},
]
# 报告为阳性或阴性报告 0325.aja/0325.zju/0413.xy/0320.nj2h
ybit = '0320.nj2h'
texdata = {
    'name': "韦苇",'report_id': "ZZ210874",'collect_date': "2021-04-10",'gender': "男",'age': "41",'patient_id': "1558305",'bed_id': "35",
    'hospital_id': "广西医科大一附院",'department_id': "呼吸内二科",'doctor_name': "唐海娟",'detect_date': "2021-04-9",'report_date': "2021-04-12",
    'sample_type': "肺泡灌洗液",'sample_volume': "10ml",'chief_complaint': "反复发热半月，咳嗽咳痰3天",'clinical_diagnosis': "-",
    'pathogen_tip': "分枝杆菌，关注结核",'drug_list': "美罗培南，莫西沙星",'is_drug_used': "是",'wbc': "16.79",'wbc': "16.79",'pmn': "-",
    'lym': "1.45",'platelet': "-",'crp': ">192",'pct': "0.52",'culture': "-",'identification': "-",'scopy': "-",'project_type': "DNA",
    'proj_type': "DNA",'report_type': "未检出明确的病原微生物",'hightyoe': hightyoe, 'lowtyoe': lowtyoe,'listData': listData,
     'results_list': results_list, 'areanet': areanet,'total_reads':'47,812,914','q30': '96.51', 'descriptions': descriptions, 'papers': papers,
    'amr': [{'species':'大肠埃希菌','species_e':'Escherichia coli', 'area':[{'mechanisms':'抗生素靶点保护', 'gene':'tet32', 'count':'212', 'coverage':'99.3%', 'drug':'四环素'}]}],
    'desc1': '基因覆盖图', 'desc2': '基因覆盖图', 'desc3': '基因覆盖图', 'ybit': ybit,
    
}


path_dir = os.getcwd()
loader = FileSystemLoader(searchpath=path_dir)
env = Environment(loader=loader)
template = env.get_template("0320_nj2h_stencil.html") # 模板文件 highSpecial
buf = template.render(texdata)
with open(os.path.join(path_dir, "nmgs1.html"), "w", encoding="utf-8") as fp:
  fp.write(buf)
