from collections import defaultdict
from os.path import isdir
from typing import DefaultDict
import xlrd,subprocess
#from docxtpl import DocxTemplate, RichText
import sys,pathlib,json,re
from multiprocessing import Pool, set_start_method

pp = pathlib.Path(sys.argv[0])
tpl_path = pp.absolute().parent / 'tpl'

infoName = ['xh','kd','report_id','sample_id','hospital','name','gender','age','patient_id','bed_id','tel','hospital_id','department_id','doctor_name','detect_date',\
    'collect_date','jc_date','report_date','proj_type','sample_type','sample_volume','sample_remained','chief_complaint','clinical_diagnosis','pathogen_tip',\
    'drug_list','is_drug_used','wbc','lym','crp','pct','pmn','platelet','culture','identification','scopy']
rgi = { 'acridine dye': "二氯基吖啶", "aminocoumarin antibiotic": "氨基香豆素类", "aminoglycoside antibiotic": "氨基糖苷类", "antibacterial free fatty acids": "FFA抗菌素游离脂肪酸", \
    "benzalkonium chloride": "苯扎氯铵", "bicyclomycin": "双环霉素", "carbapenem": "碳青霉烯类", "cephalosporin": "头孢菌素", "cephamycin": "头霉素类", \
    "cycloserine": "环丝氨酸", "diaminopyrimidine antibiotic": "促生长类", "diarylquinoline antibiotic": "二芳基喹啉", "elfamycin antibiotic": "elfamycin类", \
    "ethionamide": "乙硫异烟胺", "fluoroquinolone antibiotic": "氟喹诺酮类", "fosfomycin": "磷霉素", "fosmidomycin": "膦胺霉素", "fusidic acid": "夫西地酸", \
    "glycopeptide antibiotic": "糖肽类", "glycylcycline": "甘氨酰胺四环素", "isoniazid": "异烟肼", "lincosamide antibiotic": "林可霉素类", "macrocyclic antibiotic": "大环类", \
    'macrolide antibiotic': "大环内酯类", "monobactam": "单内酰环类", "mupirocin": "莫匹罗星", "nitrofuran antibiotic": "硝基呋喃", "nitroimidazole antibiotic": "硝基咪唑", \
    "nucleoside antibiotic": "核苷类抗生素", "nybomycin": "尼博霉素", "organoarsenic antibiotic": "有机砷", "oxazolidinone antibiotic": "噁唑烷酮抗生素", "pactamycin": "约霉素", \
    "para-aminosalicylic acid": "对氨水杨酸", "penam": "青霉烷类", "penem": "青霉烯类", "peptide antibiotic": "肽抗生素类", "phenicol antibiotic": "苯丙醇", \
    "pleuromutilin antibiotic": "截短侧耳素类", "polyamine antibiotic": "多胺", "prothionamide": "丙硫异烟胺", "pyrazinamide": "吡嗪酰胺", "rhodamine": "罗丹明", \
    "rifamycin antibiotic": "利福霉素", "streptogramin antibiotic": "链阳霉素类", "sulfonamide antibiotic": "磺胺类", "sulfone antibiotic": "砜类抗生素", \
    "tetracycline antibiotic": "四环素", "triclosan": "三氯生类", "penicillin": "青霉素类", "chloramphenic": "氯霉素类", "Minocycline": "米诺环素"}
med = {'antibiotic efflux':'抗生素外排', 'antibiotic target alteration':'抗生素靶点改变', 'antibiotic inactivation':'抗生素灭活', 'reduced permeability to antibiotic':'抗生素渗透性降低',\
    'antibiotic target protection':'抗生素靶点保护','antibiotic target replacement':'抗生素靶点置换'}
en2zn = {'bacteria':'细菌', 'virus':'病毒', 'fungi':'真菌', 'parasite':'寄生虫', 'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）'}


possample = []
negsample = []
rgisample = []
hysample = []
mzsample = []
boaosample = []
rysample = []
xysample = []
njsample = []
report_date = {}
report_sample = []
allsample = []
the_type = defaultdict(str)
sa_type = defaultdict(str)
reportid = defaultdict(str)
library = defaultdict(set)

def getSampleInfo(inputfile,inputdir):
    sn = 1
    num = 6
    samplesn = defaultdict(int)
    samplenum = defaultdict(int)
    book = xlrd.open_workbook(inputfile)
    
    sample = defaultdict(dict)

    ##读取BASIC sheet
    infosheet = book.sheet_by_index(0)
    for i in range(1, infosheet.nrows):
        sample_id = infosheet.row(i)[3].value.strip()
        allsample.append(sample_id)
        report_date[sample_id] = infosheet.row(i)[17].value.strip()
        the_type[sample_id] = infosheet.row(i)[18].value.strip()
        sample[sample_id].update({ 'the_type':the_type[sample_id] })
        sa_type[sample_id] = infosheet.row(i)[19].value.strip()
        reportid[sample_id] = infosheet.row(i)[2].value.strip()
        samplesn[sample_id] = sn
        samplenum[sample_id] = num
        for m, n in zip(infosheet.row(i), infoName):
            value = str(m.value).strip()
            if not value:
                value = '-'
            sample[sample_id].update({ n: value })
    
    ##读取模版信息
    tplsheet = book.sheet_by_index(1)
    for i in range(1, tplsheet.nrows):
        tpl = [str(j.value).strip() for j in tplsheet.row(i)]
        sample[tpl[0]].update({ 'tpl':tpl[-1] })
        if tpl[1] == '阳性':
            possample.append(tpl[0])
        else:
            negsample.append(tpl[0])
        if tpl[0] in allsample:
            report_sample.append(tpl[0])
            if tpl[-1].find('hy') > -1:
                print(f'华银报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0201.hy' ,'shuiyin':'','qmgzbit':True  })                
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                hysample.append(tpl[0])
            elif tpl[-1].find('mz') > -1:
                print(f'明志报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0201.mz' ,'shuiyin':'','qmgzbit':True })                
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                mzsample.append(tpl[0])
            elif tpl[-1].find('boao') > -1:
                print(f'博奥报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0201.boao' ,'shuiyin':'','qmgzbit':True })                
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                boaosample.append(tpl[0])
            elif tpl[-1].find('ry') > -1:
                print(f'锐翌报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0420.ry' ,'shuiyin':'','qmgzbit':True })                
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                rysample.append(tpl[0])
            elif tpl[-1].find('xy') > -1:
                print(f'湘雅报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0525.xy' ,'shuiyin':'','qmgzbit':True })
                sample[tpl[0]].update({ 'project_type':tpl[2].replace('免费','') })
                xysample.append(tpl[0])
            elif tpl[-1].find('ql') > -1:
                print(f'齐鲁报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0618.ql' ,'shuiyin':'','qmgzbit':True })
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                njsample.append(tpl[0])
            elif tpl[-1].find('cd') > -1:
                print(f'常德报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0618.cd' ,'shuiyin':'','qmgzbit':True })
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                xysample.append(tpl[0])
            elif tpl[-1].find('zk') > -1:
                print(f'中科报告：{tpl[0]}')
                sample[tpl[0]].update({ 'ybit':'0625.zk' ,'shuiyin':'','qmgzbit':True })
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                njsample.append(tpl[0])
            else:
                print(f'aja/zju/nj 报告：{tpl[0]}')
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                if tpl[-1].find('nj2h') > -1:
                    njsample.append(tpl[0])
                    sample[tpl[0]].update({ 'ybit':'0320.nj2h' ,'shuiyin':'','qmgzbit':True })
                elif tpl[-1].find('fzch') > -1:
                    njsample.append(tpl[0])
                    sample[tpl[0]].update({ 'ybit':'0201.fzch' ,'shuiyin':'','qmgzbit':True })
                elif tpl[-1].find('zju') > -1:
                    rgisample.append(tpl[0])
                    sample[tpl[0]].update({ 'ybit':'0325.zju' })
                    if tpl[-1].find('positive2') > -1 or tpl[-1].find('negative2') > -1:
                        sample[tpl[0]].update({'shuiyin':'免费测试','the_detect':'','the_signature':'','the_report_date':'','date':'','beizhu':'','qmgzbit':False})
                    else:
                        if str(tpl[-1]).find('pay') > -1:
                            sample[tpl[0]].update({'qmgzbit':True,'shuiyin':'','the_detect':'检测者：','the_signature':'审核签字：','the_report_date':'报告日期：','date':report_date[tpl[0]],'beizhu':'备注：此报告仅对本次送检样本负责！结果仅供医生参考。\n若对检测结果有疑问，请于收到报告后7个工作日内与我们联系，谢谢合作！'})
                        else:
                            sample[tpl[0]].update({'qmgzbit':True,'shuiyin':'免费测试','the_detect':'检测者：','the_signature':'审核签字：','the_report_date':'报告日期：','date':report_date[tpl[0]],'beizhu':'备注：此报告仅对本次送检样本负责！结果仅供医生参考。\n若对检测结果有疑问，请于收到报告后7个工作日内与我们联系，谢谢合作！'})
                elif tpl[-1].find('aja') > -1:
                    rgisample.append(tpl[0])
                    sample[tpl[0]].update({ 'ybit':'0325.aja' })
                    if tpl[-1].find('pay') > -1:
                        sample[tpl[0]].update({'qmgzbit':True,'shuiyin':''})
                    else:
                        sample[tpl[0]].update({'qmgzbit':True,'shuiyin':'免费测试'})

    ##绘制病原体覆盖度图
    kkplotsheet = book.sheet_by_index(6)
    the_kk = defaultdict(list)
    plot = defaultdict(list)
    for i in range(1,kkplotsheet.nrows):
        kkplot = [str(j.value).strip() for j in kkplotsheet.row(i)]
        kksa_id = kkplot[0].strip().split('-')[0]
        kkmicro = kkplot[1].replace('_',' ')
        plot[kkmicro] = kkplot[1]
        the_kk[kksa_id].append(kkmicro)
        script = f'''/home/yong_sun/bin/plot/venv/bin/python /home/yong_sun/bin/plot/plotDepthCoverage4flask.py {kkplot[0]} {kkplot[1]} /data/mngsSYS/b/reportTMP/plot'''
        subprocess.run(script, shell=True, stderr=subprocess.PIPE)

    ##读取阳性病原体信息
    possheet = book.sheet_by_index(2)
    highBacteria,highVirus,highFungi,highParasite,highSpecial = defaultdict(dict),defaultdict(dict),defaultdict(dict),defaultdict(dict),defaultdict(dict)
    lowBacteria,lowVirus,lowFungi,lowParasite,lowSpecial = defaultdict(dict),defaultdict(dict),defaultdict(dict),defaultdict(dict),defaultdict(dict)
    virusList,bacteriaList,fungiList,parasiteList,specialList,mycoList,zytList = defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list)
    description,papers = defaultdict(list),defaultdict(list)
    lowtyoe,hightyoe = defaultdict(list),defaultdict(list)
    fungi_parasiteList,bacteria_specialList = defaultdict(list),defaultdict(list)
    
    sampleposReads = defaultdict(lambda: defaultdict(int))
    sampleposInfos = defaultdict(lambda: defaultdict(list))
    allmicro = defaultdict(lambda: defaultdict(str))
    micro = defaultdict(list)
    all_micro = defaultdict(str)
        
    for i in range(1,possheet.nrows):
        pos = [str(j.value).strip() for j in possheet.row(i)]
        if pos[0] in possample:
            sampleposReads[pos[0]][pos[8]] = int(float(pos[4]))
            sampleposInfos[pos[0]][pos[8]] = pos
    
    area = defaultdict(lambda: defaultdict(list))
    high_area,low_area = defaultdict(lambda: defaultdict(list)),defaultdict(lambda: defaultdict(list))
    znen,high_znen,low_znen = defaultdict(lambda: defaultdict(list)),defaultdict(lambda: defaultdict(list)),defaultdict(lambda: defaultdict(list))
    species_type,genus_type = defaultdict(lambda: defaultdict(str)),defaultdict(lambda: defaultdict(str))
    gcount = defaultdict(lambda: defaultdict(set))
    spen2cn = defaultdict(str)
    for i in sampleposReads:
        for s, v in sorted(sampleposReads[i].items(), key=lambda x: x[1], reverse=True):
            pos = sampleposInfos[i][s]   
            species_type[pos[0]][pos[10]] = pos[2]
            genus_type[pos[0]][pos[10]] = pos[9]
            spen2cn[pos[3]] = pos[8]
            if pos[0] in rysample:
                micro[pos[0]].append(pos[8] + '( ' + pos[3] + ' )')
            abu_raw = float(pos[6])
            abu_clean = str(float('%.3f' % float(abu_raw))) if abu_raw > 0.001 else str('&lt;' + '0.001')
            gcount[pos[0]][pos[10]].add(int(float(pos[5])))         
            e_sp = {'type': pos[-3],
                    'species': pos[8],
                    'species_e': pos[3],
                    's_zn': pos[8],
                    's_en': pos[3],
                    'scount': format(int(float(pos[4])),','),
                    'abundance': str(abu_clean) + str('%'),
                    'focus': pos[7]}
            znen[pos[0]][pos[10]] = {'genus':pos[9], 'genus_e':pos[10], 'gcount': format(int(float(pos[5])), ','), 'g_en': pos[10], 'g_zn': pos[9]}
            if pos[10] in znen[pos[0]]:
                area[pos[0]][pos[10]].append(e_sp) 
            else:
                area[pos[0]][pos[10]] = [e_sp]
            if pos[7] == '高':
                high_znen[pos[0]][pos[2]] = { 'name': en2zn[pos[2]]}
                high_sp = { 'species_cn': pos[8], 'species_e': pos[3] }
            else:
                low_znen[pos[0]][pos[2]] = { 'name': en2zn[pos[2]]}
                low_sp = { 'species_cn': pos[8], 'species_e': pos[3] }
            if pos[7] == '高':
                if pos[2] in high_znen[pos[0]]:
                    high_area[pos[0]][pos[2]].append(high_sp)
                else:
                    high_area[pos[0]][pos[2]] = [high_sp] 
            else:
                if pos[2] in low_znen[pos[0]]:
                    low_area[pos[0]][pos[2]].append(low_sp)
                else:
                    low_area[pos[0]][pos[2]] = [low_sp] 
    
    for k,v in micro.items():
        all_micro[k] = '、'.join(i for i in micro[k])

    listData = defaultdict(list)
    for i in area:
        if i not in rysample:
            for s, v in area[i].items():
                znen[i][s]['area'] = area[i][s]
                if species_type[i][s] == 'fungi':
                    if len(gcount[i][s]) == 1: 
                        fungiList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        fungiList[i].append(znen[i][s])
                elif species_type[i][s] == 'bacteria':
                    if len(gcount[i][s]) == 1:
                        bacteriaList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        bacteriaList[i].append(znen[i][s])
                elif species_type[i][s] == 'virus':
                    if len(gcount[i][s]) == 1:
                        virusList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        virusList[i].append(znen[i][s])
                elif species_type[i][s] == 'parasite':
                    if len(gcount[i][s]) == 1:
                        parasiteList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        parasiteList[i].append(znen[i][s])
                elif species_type[i][s] == 'special':
                    if len(gcount[i][s]) == 1:
                        specialList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        specialList[i].append(znen[i][s])
                else:
                    print(f'{i}物种类型单词写错了！')
                fungi_parasiteList[i] = fungiList[i] + parasiteList[i]
                bacteria_specialList[i] = bacteriaList[i] + specialList[i]
        else:
            for s, v in area[i].items():
                znen[i][s]['area'] = area[i][s]
                if species_type[i][s] == 'fungi':  
                    if len(gcount[i][s]) == 1: 
                        fungiList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        fungiList[i].append(znen[i][s])
                elif species_type[i][s] == 'bacteria':
                    if len(gcount[i][s]) == 1:
                        bacteriaList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        bacteriaList[i].append(znen[i][s])
                elif species_type[i][s] == 'virus':
                    if len(gcount[i][s]) == 1:
                        virusList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        virusList[i].append(znen[i][s])
                elif species_type[i][s] == 'parasite':
                    if len(gcount[i][s]) == 1:
                        parasiteList[i].append(znen[i][s])
                    else:
                        znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                        parasiteList[i].append(znen[i][s])
                elif species_type[i][s] == 'special':
                    if genus_type[i][s] == '分枝杆菌属':
                        if len(gcount[i][s]) == 1:
                            mycoList[i].append(znen[i][s])
                        else:
                            znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                            mycoList[i].append(znen[i][s])
                    else:
                        if len(gcount[i][s]) == 1:
                            zytList[i].append(znen[i][s])
                        else:
                            znen[i][s]['gcount'] = sum(list(gcount[i][s]))
                            zytList[i].append(znen[i][s])
                else:
                    print(f'{i}物种类型单词写错了！')  

    for i in high_area:
        if i not in rysample:
            for s, v in high_area[i].items():
                high_znen[i][s]['species'] = high_area[i][s]
                if s == 'fungi':
                    highFungi[i].update(high_znen[i][s])
                elif s == 'bacteria':
                    highBacteria[i].update(high_znen[i][s])
                elif s == 'virus':
                    highVirus[i].update(high_znen[i][s])
                elif s == 'parasite':
                    highParasite[i].update(high_znen[i][s])
                elif s == 'special':
                    highSpecial[i].update(high_znen[i][s])
                else:
                    print(f'{i}物种类型单词写错了！')

    for i in low_area:
        if i not in rysample:
            for s, v in low_area[i].items():
                low_znen[i][s]['species'] = low_area[i][s]
                if s == 'fungi':
                    lowFungi[i].update(low_znen[i][s])
                elif s == 'bacteria':
                    lowBacteria[i].update(low_znen[i][s])
                elif s == 'virus':
                    lowVirus[i].update(low_znen[i][s])
                elif s == 'parasite':
                    lowParasite[i].update(low_znen[i][s])
                elif s == 'special':
                    lowSpecial[i].update(low_znen[i][s])
                else:
                    print(f'{i}物种类型单词写错了！')

    def deal(the_id):
        if the_id in boaosample:
            hightyoe[the_id] = [highBacteria[the_id],highVirus[the_id],highFungi[the_id],highParasite[the_id],highSpecial[the_id]]
            lowtyoe[the_id] = [lowBacteria[the_id],lowVirus[the_id],lowFungi[the_id],lowParasite[the_id],lowSpecial[the_id]]
            if not highFungi[the_id] and not highBacteria[the_id] and not highVirus[the_id] and not highParasite[the_id] and not highSpecial[the_id]:
                hightyoe[the_id] = []
            if not lowFungi[the_id] and not lowBacteria[the_id] and not lowVirus[the_id] and not lowParasite[the_id] and not lowSpecial[the_id]:
                lowtyoe[the_id] = []
        else:
            if not highFungi[the_id]:
                highFungi[the_id].update({'name':'真菌', 'species':[]})
            if not highBacteria[the_id]:
                highBacteria[the_id].update({'name':'细菌', 'species':[]})
            if not highVirus[the_id]:
                highVirus[the_id].update({'name':'病毒', 'species':[]})
            if not highParasite[the_id]:
                highParasite[the_id].update({'name':'寄生虫', 'species':[]})
            if not highSpecial[the_id]:
                highSpecial[the_id].update({'name':'特殊病原体（包括分枝杆菌、支原体/衣原体等）', 'species':[]})
            hightyoe[the_id] = [highBacteria[the_id],highVirus[the_id],highFungi[the_id],highParasite[the_id],highSpecial[the_id]]
            if not lowFungi[the_id]:
                lowFungi[the_id].update({'name':'真菌', 'species':[]})
            if not lowBacteria[the_id]:
                lowBacteria[the_id].update({'name':'细菌', 'species':[]})
            if not lowVirus[the_id]:
                lowVirus[the_id].update({'name':'病毒', 'species':[]})
            if not lowParasite[the_id]:
                lowParasite[the_id].update({'name':'寄生虫', 'species':[]})
            if not lowSpecial[the_id]:
                lowSpecial[the_id].update({'name':'特殊病原体（包括分枝杆菌、支原体/衣原体等）', 'species':[]})
            lowtyoe[the_id] = [lowBacteria[the_id],lowVirus[the_id],lowFungi[the_id],lowParasite[the_id],lowSpecial[the_id]]
        return hightyoe[the_id],lowtyoe[the_id]
    
    def RGI(the_id):
        flag = 0
        b = defaultdict(list)
        amr = defaultdict(lambda: defaultdict(list))
        amr_area = defaultdict(lambda: defaultdict(list))
        amr_summary = defaultdict(str)
        for k in iter(library[the_id]):
            with open(f'{inputdir}/RGI/{k}.gene_mapping_data.txt') as rgifile:
                fh = rgifile.readlines()
                if len(fh) == 1:
                    amr_summary[the_id] = '通过分析，未检出耐药基因。'
                    b[the_id] = []
                else:                          
                    for j in fh[1:]:
                        e = j.strip().split('\t')
                        if allmicro[the_id][e[-1]]:
                            flag = 1
                            mechanisms = ';'.join([med[x] for x in e[-2].split('; ')])
                            drugs = e[4].split('; ')
                            rgis = '; '.join([rgi[x] for x in drugs])
                            en = e[-1]
                            drug_en = e[4]
                            genename = e[1]
                            coverage = str(float('%.1f' % float(e[3]))) + str('%')
                            amr[the_id][e[-1]] = { 'species':spen2cn[e[-1]],'species_e': e[-1] }
                            if the_id in rysample:
                                amr[i][e[-1]] = { 'species_e':en, 'species':spen2cn[e[-1]] }
                                e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'drug_en':drug_en, 'drug_zn':rgis }
                            else:
                                if len(drugs) <= 3:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':rgis }
                                else:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':'多重耐药' }
                            if e[-1] in amr[the_id]:
                                amr_area[the_id][e[-1]].append(e_sp)
                            else:
                                amr_area[the_id][e[-1]] = [e_sp]
        if flag == 1:                    
            for k,v in amr_area[the_id].items():
                amr_summary[the_id] = '通过分析，发现患者可能对以下抗生素耐药。'
                amr[the_id][k]['area'] = amr_area[the_id][k]
                b[the_id].append(amr[the_id][k])                 
        elif flag == 0:
            amr_summary[the_id] = '通过分析，未检出耐药基因。'
            b[the_id] = []      
        return amr_summary[the_id],b[the_id]

    the_low = defaultdict(list)
    for i in sampleposReads:
        for s, v in sorted(sampleposReads[i].items(), key=lambda x: x[1], reverse=True):
            pos = sampleposInfos[i][s]
            the_low[pos[0]].append(pos[7])
            library[pos[0]].add(pos[1])
            if pos[0] not in rysample:
                microRT = pos[8] + '\n' + pos[3]
                allmicro[pos[0]][pos[3]] = microRT
            else:
                allmicro[pos[0]][pos[3]] = pos[8]
            ze = True if re.search(r'[\u4e00-\u9fa5]',pos[-1])  else False
            if pos[0] in rgisample:
                if pos[-2] != '' and pos[-1] != '':
                    if pos[3] in the_kk[pos[0]]:
                        descRT = { 'url_name':f'/data/mngsSYS/b/reportTMP/plot/{pos[1]}.{plot[pos[3]]}.png','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                        paperRT = { 'namecon':pos[-1], 'conbit':ze }
                    else:
                        descRT = { 'url_name':'-','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                        paperRT = { 'namecon':pos[-1], 'conbit':ze }
                elif pos[-2] != '' and pos[-1] == '':
                    if pos[3] in the_kk[pos[0]]:
                        descRT = { 'url_name':f'/data/mngsSYS/b/reportTMP/plot/{pos[1]}.{plot[pos[3]]}.png','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                        paperRT = { 'namecon':pos[-1], 'conbit':ze }
                    else:
                        descRT = { 'url_name':'-','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                        paperRT = { 'namecon':pos[-1], 'conbit':ze }
                else:
                    if pos[3] in the_kk[pos[0]]:
                        descRT = { 'url_name':f'/data/mngsSYS/b/reportTMP/plot/{pos[1]}.{plot[pos[3]]}.png','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                        paperRT = { 'namecon':pos[-1], 'conbit':ze }
                    else:
                        descRT = { 'url_name':'-','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                        paperRT = { 'namecon':pos[-1], 'conbit':ze }
                papers[pos[0]].append(paperRT)
                description[pos[0]].append(descRT)
            else:
                if pos[-2] != '' and pos[-1] != '':
                    descRT = { 'url_name':'-','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                    paperRT = { 'namecon':pos[-1], 'conbit':ze }
                elif pos[-2] != '' and pos[-1] == '':
                    descRT = { 'url_name':'-','name': pos[8], 'english':pos[3], 'explain':pos[-2] }
                    paperRT = { 'namecon':'NA', 'conbit':ze }
                else:
                    descRT = { 'url_name':'-','name': pos[8], 'english':pos[3], 'explain':'NA' }
                    paperRT = { 'namecon':'NA', 'conbit':ze }
                papers[pos[0]].append(paperRT)
                description[pos[0]].append(descRT)
            samplesn[pos[0]] += 1
            samplenum[pos[0]] += 1     

    ##读取阴性疑似病原体和背景微生物信息
    negsheet = book.sheet_by_index(3)
    backlist = defaultdict(list)
    all_back = defaultdict(list)
    all_backlist = defaultdict(str)
    sampleneg_bj_Reads = defaultdict(lambda: defaultdict(int))
    sampleneg_bj_Infos = defaultdict(lambda: defaultdict(list))
    
    for i in range(1,negsheet.nrows):
        neg = [str(j.value).strip() for j in negsheet.row(i)]
        if neg[5] == '疑似背景微生物':
            sampleneg_bj_Reads[neg[0]][neg[2]] = int(float(neg[4]))
            sampleneg_bj_Infos[neg[0]][neg[2]] = neg
    
    for i in sampleneg_bj_Reads:
        id = i.strip().split('\t')[0]
        if id in boaosample:
            with open(f'{sys.argv[4]}/{reportid[id]}_背景列表.xls','w',encoding='gbk') as boaoback:
                boaoback.write(f'name\tChinese\thit_reads\n')
                for s, v in sorted(sampleneg_bj_Reads[i].items(), key=lambda x: x[1], reverse=True):
                    neg = sampleneg_bj_Infos[i][s]
                    if neg[0] == str(id):
                        boaoback.write(f'{neg[2]}\t{neg[3]}\t{int(float(neg[4]))}\n')
                boaoback.close()
        else:                    
            for s, v in sorted(sampleneg_bj_Reads[i].items(), key=lambda x: x[1], reverse=True):
                neg = sampleneg_bj_Infos[i][s]
                ze = True if re.search(r'[\u4e00-\u9fa5]',neg[7])  else False
                if neg[0] in rysample:
                    all_back[neg[0]].append(neg[3] + '( ' + neg[2] + ' )')
                if neg[7] != '':
                    descRT = { 'url_name':'-','name': neg[3], 'english':neg[2], 'explain':neg[6] }
                    paperRT = { 'namecon':neg[7], 'conbit':ze }
                    papers[neg[0]].append(paperRT)
                    description[neg[0]].append(descRT)
                if neg[0] not in rysample:
                    backlist[neg[0]].append({'type':neg[1],'microbe':neg[3],'microbe_e':neg[2], 'count':f'{int(float(neg[4])):,}','note':neg[5]})
                else:
                    backlist[neg[0]].append({ 'g_en':neg[-2],'g_zn':neg[-3],'g_count':f'{int(float(neg[-1])):,}', 'type':neg[1], 'microbe':neg[3],'microbe_e':neg[2], 'zn':neg[3],'en':neg[2],'count':f'{int(float(neg[4])):,}','note':'疑似病原体，人体共生条件致病菌' })
            for k,v in all_back.items():
                all_backlist[k] = '、'.join(i for i in all_back[k])
            
    ###模版内容添加   
    for i in sample:
        if i in possample:
            summary = defaultdict(str)
            amr_b = defaultdict(list)
            if i in rgisample or i in hysample or i in xysample or i in njsample:
                listData[i] = [bacteriaList[i], virusList[i], fungiList[i], parasiteList[i], specialList[i]]
                hightyoe[i],lowtyoe[i] = deal(i)
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':'检出以下疑似病原体'})
                sample[i].update({ 'the_desc':'(4) 带*标记病原体表示，该病原体低于检测阈值，需要临床综合考虑其致病可能性。'}) if '低*' in the_low[i] else sample[i].update({ 'the_desc':''})
                sample[i].update({ 'hightyoe': hightyoe[i], 'lowtyoe': lowtyoe[i] })
                sample[i].update({ 'listData': listData[i] })
                sample[i].update({ 'areanet':backlist[i] }) if backlist[i] else sample[i].update({ 'areanet':[] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
                summary[i],amr_b[i] = RGI(i)
                sample[i].update({'amr_summary':summary[i]})
                sample[i].update({ 'amr':amr_b[i] })
            elif i in boaosample:
                hightyoe[i],lowtyoe[i] = deal(i)
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':'检出以上疑似病原体'})
                sample[i].update({ 'hightyoe': hightyoe[i], 'lowtyoe': lowtyoe[i] })
                sample[i].update({ 'bacteria_specialList':bacteria_specialList[i] }) if bacteria_specialList[i] else sample[i].update({ 'bacteria_specialList':[]})
                sample[i].update({ 'virusList':virusList[i] }) if virusList[i] else sample[i].update({ 'virusList':[]})
                sample[i].update({ 'fungi_parasiteList':fungi_parasiteList[i] }) if fungi_parasiteList[i] else sample[i].update({ 'fungi_parasiteList':[]})
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                summary[i],amr_b[i] = RGI(i)
                sample[i].update({'amr_summary':summary[i]})
                sample[i].update({ 'amr':amr_b[i] })
            elif i in rysample:
                sample[i].update({ 'bacteriaList':bacteriaList[i] }) if bacteriaList[i] else sample[i].update({ 'bacteriaList':[]})
                sample[i].update({ 'fungiList':fungiList[i]}) if fungiList[i] else sample[i].update({ 'fungiList':[]})
                sample[i].update({ 'virusList':virusList[i]}) if virusList[i] else sample[i].update({ 'virusList':[]})
                sample[i].update({ 'parasiteList':parasiteList[i]}) if parasiteList[i] else sample[i].update({ 'parasiteList':[]})
                sample[i].update({ 'mycoList':mycoList[i]}) if mycoList[i] else sample[i].update({ 'mycoList':[]})
                sample[i].update({ 'zytList':zytList[i]}) if zytList[i] else sample[i].update({ 'zytList':[]})
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':f'该样本中检测到的病原体有{all_micro[i]}' })
                sample[i].update({ 'all_backlist':all_backlist[i] })
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[] })
                summary[i],amr_b[i] = RGI(i)
                sample[i].update({'amr_summary':summary[i]})
                sample[i].update({ 'amr':amr_b[i] })
                #b = []
                #amr = defaultdict(lambda: defaultdict(list))
                #amr_area = defaultdict(lambda: defaultdict(list))
                #flag = 0
                #for k in iter(library[i]):
                #    with open(f'{inputdir}/RGI/{k}.gene_mapping_data.txt') as rgifile:
                #        fh = rgifile.readlines()
                #        if len(fh) == 1:
                #            sample[i].update({ 'amr':[] })
                #        else:
                #            for j in fh[1:]:
                #                e = j.strip().split('\t')
                #                if allmicro[i][str(e[-1])]:
                #                    flag = 1
                #                    mechanisms = ';'.join([med[x] for x in e[-2].split('; ')])
                #                    en = e[-1]
                #                    drug_en = e[4]
                #                    drugs = e[4].split('; ')
                #                    rgis = '; '.join([rgi[x] for x in drugs])
                #                    genename = e[1]
                #                    coverage = str(float('%.1f' % float(e[3]))) + str('%')
                #                    amr[i][e[-1]] = { 'species_e':en, 'species':spen2cn[e[-1]] }
                #                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'drug_en':drug_en, 'drug_zn':rgis }
                #                    if e[-1] in amr[i]:
                #                        amr_area[i][e[-1]].append(e_sp)
                #                    else:
                #                        amr_area[i][e[-1]] = [e_sp]
                #if flag == 1:
                #    for k,v in amr_area[i].items():
                #        amr[i][k]['area'] = amr_area[i][k]
                #        b.append(amr[i][k])
                #    sample[i].update({ 'amr':b })                    
                #elif flag == 0:
                #    amr = []
            else:
                listData[i] = [bacteriaList[i], virusList[i], fungiList[i], parasiteList[i], specialList[i]]
                hightyoe[i],lowtyoe[i] = deal(i)
                sample[i].update({ 'hightyoe': hightyoe[i], 'lowtyoe': lowtyoe[i] })
                sample[i].update({ 'listData': listData[i] })
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':'检出以下疑似病原体'})
                sample[i].update({ 'areanet':backlist[i] }) if backlist[i] else sample[i].update({ 'areanet':[] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
                summary[i],amr_b[i] = RGI(i)
                sample[i].update({'amr_summary':summary[i]})
                sample[i].update({ 'amr':amr_b[i] })                           
        elif i in negsample:
            if i in rgisample or i in hysample or i in xysample or i in njsample:
                listData[i] = [bacteriaList[i], virusList[i], fungiList[i], parasiteList[i], specialList[i]]
                hightyoe[i],lowtyoe[i] = deal(i)
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':'未检出明确的病原微生物'})
                sample[i].update({ 'hightyoe': hightyoe[i], 'lowtyoe': lowtyoe[i] })
                sample[i].update({ 'listData': listData[i] })
                sample[i].update({ 'areanet':backlist[i] }) if backlist[i] else sample[i].update({ 'areanet':[] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
                sample[i].update({ 'amr_summary':'通过分析，未检出耐药基因。' })
                sample[i].update({ 'amr':[] })
            elif i in boaosample:
                hightyoe[i],lowtyoe[i] = deal(i)
                sample[i].update({ 'hightyoe': hightyoe[i], 'lowtyoe': lowtyoe[i] })
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':'未检出明确的病原微生物'})
                sample[i].update({ 'bacteria_specialList':[]})
                sample[i].update({ 'virusList':[]})
                sample[i].update({ 'fungi_parasiteList':[]})
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'amr':[] })
            elif i in rysample:
                sample[i].update({ 'amr':[] })
                sample[i].update({ 'bacteriaList':[]})
                sample[i].update({ 'fungiList':[]})
                sample[i].update({ 'virusList':[]})
                sample[i].update({ 'parasiteList':[]})
                sample[i].update({ 'mycoList':[]})
                sample[i].update({ 'zytList':[]})
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':'该样本中未检测到病原体' })
                sample[i].update({ 'all_backlist':all_backlist[i] })
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[] })
            else:
                sample[i].update({ 'amr':[] })
                sample[i].update({ 'results_list':['细菌', '病毒', '真菌', '寄生虫', '特殊病原体（包括分枝杆菌、支原体/衣原体等）'] })
                sample[i].update({ 'report_type':'未检出明确的病原微生物'})
                hightyoe[i],lowtyoe[i] = deal(i)
                sample[i].update({ 'hightyoe': hightyoe[i], 'lowtyoe': lowtyoe[i] })
                sample[i].update({ 'areanet':backlist[i] }) if backlist[i] else sample[i].update({ 'areanet':[] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
    
    ##读取数据量信息
    runstatsheet = book.sheet_by_index(5)
    total_reads,human_reads,nonhuman_reads,micro_reads,nonhuman_rate,q30 = defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(float),defaultdict(float)
    CSF_DNA_total,CSF_DNA_the_num,CSF_DNA_human,CSF_DNA_nonhuman,CSF_DNA_micro,CSF_DNA_q30 = defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(float)
    CSF_CF_total,CSF_CF_the_num,CSF_CF_human,CSF_CF_nonhuman,CSF_CF_micro,CSF_CF_q30 = defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(float)
    CSF_RNA_total,CSF_RNA_the_num,CSF_RNA_human,CSF_RNA_nonhuman,CSF_RNA_micro,CSF_RNA_q30 = defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(float)
    the_num = defaultdict(int)
    the_sa = set()
    for i in range(1,runstatsheet.nrows):
        stat = [str(j.value).strip() for j in runstatsheet.row(i)]
        sa_id = stat[0].strip().split('-')
        the_sa.add(sa_id[0])
        if sa_type[sa_id[0]] == '脑脊液':
            if len(sa_id) == 1 or sa_id[-1] == '1':
                CSF_DNA_the_num[sa_id[0]] = 1
                CSF_DNA_total[sa_id[0]] = int(float(stat[1]))
                CSF_DNA_human[sa_id[0]] = int(float(stat[2]))
                CSF_DNA_nonhuman[sa_id[0]] = int(float(stat[3]))
                CSF_DNA_micro[sa_id[0]] = int(float(stat[4]))
                CSF_DNA_q30[sa_id[0]] = float(stat[5])
            elif len(sa_id) > 1 and sa_id[-1] == 'CF':
                CSF_CF_the_num[sa_id[0]] = 1
                CSF_CF_total[sa_id[0]] = int(float(stat[1]))
                CSF_CF_human[sa_id[0]] = int(float(stat[2]))
                CSF_CF_nonhuman[sa_id[0]] = int(float(stat[3]))
                CSF_CF_micro[sa_id[0]] = int(float(stat[4]))
                CSF_CF_q30[sa_id[0]] = float(stat[5])
            elif len(sa_id) > 1 and sa_id[-1] == 'R':
                CSF_RNA_the_num[sa_id[0]] = 1
                CSF_RNA_total[sa_id[0]] = int(float(stat[1]))
                CSF_RNA_human[sa_id[0]] = int(float(stat[2]))
                CSF_RNA_nonhuman[sa_id[0]] = int(float(stat[3]))
                CSF_RNA_micro[sa_id[0]] = int(float(stat[4]))
                CSF_RNA_q30[sa_id[0]] = float(stat[5])        
        else:
            if sa_id[0] not in total_reads:
                the_num[sa_id[0]] = 1
                total_reads[sa_id[0]] = int(float(stat[1]))
                human_reads[sa_id[0]] = int(float(stat[2]))
                nonhuman_reads[sa_id[0]] = int(float(stat[3]))
                micro_reads[sa_id[0]] = int(float(stat[4]))
                q30[sa_id[0]] = float(stat[5])
            else:
                the_num[sa_id[0]] += 1
                total_reads[sa_id[0]] += int(float(stat[1]))
                human_reads[sa_id[0]] += int(float(stat[2]))
                nonhuman_reads[sa_id[0]] += int(float(stat[3]))
                micro_reads[sa_id[0]] += int(float(stat[4]))
                q30[sa_id[0]] += float(stat[5])

    for k in iter(the_sa):
        if sa_type[k] == '脑脊液':
            if CSF_CF_total[k] > 15_000_000:
                the_num[k] = CSF_CF_the_num[k] + CSF_RNA_the_num[k]
                total_reads[k] = CSF_CF_total[k] + CSF_RNA_total[k]
                human_reads[k] = CSF_CF_human[k] + CSF_RNA_human[k]
                nonhuman_reads[k] = CSF_CF_nonhuman[k] + CSF_RNA_nonhuman[k]
                micro_reads[k] = CSF_CF_micro[k] + CSF_RNA_micro[k]
                q30[k] = CSF_CF_q30[k] + CSF_RNA_q30[k]
            elif CSF_DNA_total[k] > 15_000_000 and CSF_CF_total[k] < 15_000_000:
                the_num[k] = CSF_DNA_the_num[k] + CSF_RNA_the_num[k]
                total_reads[k] = CSF_DNA_total[k] + CSF_RNA_total[k]
                human_reads[k] = CSF_DNA_human[k] + CSF_RNA_human[k]
                nonhuman_reads[k] = CSF_DNA_nonhuman[k] + CSF_RNA_nonhuman[k]
                micro_reads[k] = CSF_DNA_micro[k] + CSF_RNA_micro[k]
                q30[k] = CSF_DNA_q30[k] + CSF_RNA_q30[k]
            elif CSF_DNA_total[k] < 15_000_000 and CSF_CF_total[k] < 15_000_000:
                the_num[k] = CSF_CF_the_num[k] + CSF_RNA_the_num[k]
                total_reads[k] = CSF_CF_total[k] + CSF_RNA_total[k]
                human_reads[k] = CSF_CF_human[k] + CSF_RNA_human[k]
                nonhuman_reads[k] = CSF_CF_nonhuman[k] + CSF_RNA_nonhuman[k]
                micro_reads[k] = CSF_CF_micro[k] + CSF_RNA_micro[k]
                q30[k] = CSF_CF_q30[k] + CSF_RNA_q30[k]
            nonhuman_rate[k] = nonhuman_reads[k]/total_reads[k] * 100
            sample[k].update({ 'total_reads':format(total_reads[k],','), 'human_reads':format(human_reads[k],','), \
                           'nonhuman_reads':format(nonhuman_reads[k],','),'micro_reads':format(micro_reads[k],','), \
                           'nonhuman_rate':str('%.2f' % (float(nonhuman_rate[k]))),'q30':str('%.2f' % float((q30[k]/int(the_num[k])))) })
        else:
            nonhuman_rate[k] = nonhuman_reads[k]/total_reads[k] * 100
            sample[k].update({ 'total_reads':format(total_reads[k],','), 'human_reads':format(human_reads[k],','), \
                           'nonhuman_reads':format(nonhuman_reads[k],','),'micro_reads':format(micro_reads[k],','), \
                           'nonhuman_rate':str('%.2f' % (float(nonhuman_rate[k]))),'q30':str('%.2f' % float((q30[k]/int(the_num[k])))) })
    return sample

def main():
    outdir = sys.argv[4]
    jobid = sys.argv[3]
    sample = getSampleInfo(f'{sys.argv[1]}',f'{sys.argv[2]}')
    for k, v in sample.items():
        if k in report_sample:
            with open(f'/data/mngsSYS/b/reportTMP/1/json/{jobid}_{k}.json', 'w', encoding='utf-8') as out:
                json.dump(v, out,ensure_ascii = False)
            out.close()
            script = f'''/home/runmngs/test/html2pdf/venv/bin/python /data/softwares/mngs_scripts/report/html2pdf/html2pdf.py /data/mngsSYS/b/reportTMP/1/json/{jobid}_{k}.json {outdir}'''
            subprocess.run(script,shell=True)

if __name__ == '__main__':
    main()
