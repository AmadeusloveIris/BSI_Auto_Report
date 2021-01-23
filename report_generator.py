import time
import jinja2
from os import path
import pandas as pd
from PIL import Image
from glob import glob
from docx.shared import Mm
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate, InlineImage

class Report(object):
    def __init__(self,report_path:str,order_no:str,sample_name:str,heavy_chain_rmass,light_chain_rmass):
        self.report_path = report_path
        soup = BeautifulSoup(open(path.join(report_path, 'report.html')), 'html.parser')
        section_dict = {}
        for i in soup.find_all('div', {'class': 'section'})[1:3]:
            section_dict.update({i.h2.get_text().split('_')[-1]: i})
        heavy_chain_report = section_dict['Heavy']
        light_chain_report = section_dict['Light']
        self.tpl = DocxTemplate('BSI_report_templete.docx')
        #order_no = input('Please input order number.\n')
        #sample_name = input('Please input sample name.\n')
        hsequence_format = self.get_sequence(heavy_chain_report)
        lsequence_format = self.get_sequence(light_chain_report)
        heavy_chain_cmass = self.get_chain_mass(heavy_chain_report)
        #heavy_chain_rmass = input('Please input heavy chain real mass detected by BioPharma Finder\n')
        light_chain_cmass = self.get_chain_mass(light_chain_report)
        #light_chain_rmass = input('Please input light chain real mass detected by BioPharma Finder\n')
        hpeptides_format = self.get_peptide_confidence_mapping(heavy_chain_report)
        lpeptides_format = self.get_peptide_confidence_mapping(light_chain_report)
        htIL_format = self.get_IL_loction(heavy_chain_report)
        ltIL_format = self.get_IL_loction(light_chain_report)
        hcoverage = self.get_proteinseq_confidence(heavy_chain_report, report_path)
        lcoverage = self.get_proteinseq_confidence(light_chain_report, report_path)
        hmap_format = self.get_typical_peptide_map(heavy_chain_report, report_path)
        lmap_format = self.get_typical_peptide_map(light_chain_report, report_path)
        FDR_hmapping = self.get_FDR_mapping(report_path, 'H')
        FDR_lmapping = self.get_FDR_mapping(report_path, 'L')
        self.context = {
            'Order_Number': order_no,
            'Date': str(time.strftime("%b %d %Y", time.localtime())),
            'Sample_Name': sample_name,
            'hsequence': hsequence_format,
            'heavy_chain_cmass': heavy_chain_cmass,
            'heavy_chain_rmass': heavy_chain_rmass,
            'hpeptides': hpeptides_format,
            'htIL': htIL_format,
            'lsequence': lsequence_format,
            'light_chain_cmass': light_chain_cmass,
            'light_chain_rmass': light_chain_rmass,
            'lpeptides': lpeptides_format,
            'ltIL': ltIL_format,
            'typical_hpeptide_image': InlineImage(self.tpl, hcoverage, width=Mm(165), height=Mm(79)),
            'typical_lpeptide_image': InlineImage(self.tpl, lcoverage, width=Mm(165), height=Mm(79)),
            'typical_peptide_hmap': hmap_format,
            'typical_peptide_lmap': lmap_format,
            'FDR_hmap': FDR_hmapping,
            'FDR_lmap': FDR_lmapping}
        self.jinja_env = jinja2.Environment(autoescape=True)

    def get_sequence(self,select_chain):
        sequence = ""
        for i in select_chain.table.find_all('td'):
            if not i.get_text().isdecimal():
                sequence += "".join(i.get_text().split())
        temp = [sequence[i:i + 10] for i in range(0, len(sequence), 10)]
        temp = [temp[i:i + 5] for i in range(0, len(temp), 5)]
        temp = [" ".join(i) for i in temp]
        sequence_format = [
            {'llabel': "{:0>3d}".format(50 * i + 1), 'rlabel': "{:0>3d}".format(50 * (i + 1)), 'col': temp[i]} for i in
            range(len(temp))]
        sequence_format[-1]['rlabel'] = len(sequence)
        return sequence_format

    def get_chain_mass(self, select_chain):
        return select_chain.find_all('p')[0].get_text().split()[2]

    def get_proteinseq_confidence(self, select_chain, report_path):
        img_abpath = path.join(report_path,
                               path.join('img', select_chain.find_all('div', {'class': 'coverage'})[0].img['src'][4:]))
        img = Image.open(img_abpath)
        coverage = img.crop((25, 0, 1570, 730))
        coverage.save(path.join('temp', select_chain.h2.get_text().split('_')[-1] + '_confidence.jpg'))
        return path.join('temp', select_chain.h2.get_text().split('_')[-1] + '_confidence.jpg')

    def get_peptide_confidence_mapping(self, select_chain):
        tpeptide_confidence_format = []
        temp = select_chain.find_all('table', {'class': 'peptides'})
        for tr in temp[0].find_all('tr')[1:]:
            column_flag = 0
            temp_dict = {}
            for td in tr.find_all('td'):
                if column_flag == 0:
                    if int(td.get_text().split('-')[0]) > 150:
                        break
                    else:
                        temp_dict.update({'position': td.get_text()})
                elif column_flag == 1:
                    if 'Pepsin' in td.get_text():
                        temp_dict.update({'enzyme': 'Pepsin'})
                    elif td.get_text() == 'Chymo':
                        temp_dict.update({'enzyme': 'Chymotrypsin'})
                    else:
                        temp_dict.update({'enzyme': td.get_text()})
                elif column_flag == 2:
                    temp_dict.update({'psm': td.get_text()})
                elif column_flag == 3:
                    temp_dict.update({'mass': td.get_text()})
                elif column_flag == 4:
                    temp_dict.update({'ppm': td.get_text()})
                elif column_flag == 5:
                    temp_dict.update({'abundance': "E+".join(td.get_text().split('E'))})
                elif column_flag == 6:
                    temp_dict.update({'sequence': td.get_text()})
                column_flag += 1
            if temp_dict:
                tpeptide_confidence_format.append(temp_dict)
        return tpeptide_confidence_format

    def get_FDR_mapping(self, report_path, HL_flag):
        if HL_flag == 'H':
            img = Image.open(path.join(report_path, 'hcoverage.png'))
        else:
            img = Image.open(path.join(report_path, 'lcoverage.png'))
        img = img.crop((60, 0, 1170, img.size[-1]))
        pix = img.load()
        color_cells_location = []
        for x in range(img.size[0]):
            for y in range(img.size[-1]):
                if pix[x, y] == (255, 0, 0):
                    color_cells_location.append(y)
        start_anchor = self.get_FDR_anchor(color_cells_location)
        end_anchor = start_anchor[1:] + [img.size[-1]+26]

        for i in range(len(start_anchor)):
            if start_anchor[i] + 603 < end_anchor[i]:
                if HL_flag == 'H':
                    img.crop((0, start_anchor[i] - 26, img.size[0], start_anchor[i] + 603)).save(
                        path.join('temp', '{}{}.png'.format('hfdr', i)))
                elif HL_flag == 'L':
                    img.crop((0, start_anchor[i] - 26, img.size[0], start_anchor[i] + 603)).save(
                        path.join('temp', '{}{}.png'.format('lfdr', i)))
            else:
                if HL_flag == 'H':
                    img.crop((0, start_anchor[i] - 26, img.size[0], end_anchor[i] - 26)).save(
                        path.join('temp', '{}{}.png'.format('hfdr', i)))
                elif HL_flag == 'L':
                    img.crop((0, start_anchor[i] - 26, img.size[0], end_anchor[i] - 26)).save(
                        path.join('temp', '{}{}.png'.format('lfdr', i)))

        if HL_flag == 'H':
            FDR_mapping_format = [{'img': InlineImage(self.tpl, img_name, width=Mm(165))} for img_name in
                                  sorted(glob(path.join('temp', 'hfdr*')))]
        else:
            FDR_mapping_format = [{'img': InlineImage(self.tpl, img_name, width=Mm(165))} for img_name in
                                  sorted(glob(path.join('temp', 'lfdr*')))]

        return FDR_mapping_format

    def get_typical_peptide_map(self, select_chain, report_path):
        temp = select_chain.find_all('div', {'class': "support-spectra"})[0]
        discription = []
        img = []
        stop_flag = 0
        for i in range(int(len(temp.find_all('span')) / 2)):
            if int(temp.find_all('span')[2 * i].get_text().split()[1].split('-')[0]) < 151:
                stop_flag = i
            discription.append(
                " ".join([temp.find_all('span')[2 * i].get_text(), temp.find_all('span')[2 * i + 1].get_text()]))
        for i in temp.find_all('img'):
            img.append(i.get('src')[4:])
        map_format = []
        for i in range(stop_flag):
            map_format.append({'title': discription[i],
                               'img': InlineImage(self.tpl, path.join(report_path, 'img', img[i]), width=Mm(165),
                                                  height=Mm(84.9))})
        return map_format

    def get_IL_loction(self, select_chain):
        tIL = pd.read_html(str(select_chain.find_all('div', {'class': 'il-stats subsection unbreakable'})[0].table))[0]
        temp = tIL['Position'].str.split('@', expand=True)
        tIL = pd.concat([tIL['Region'], temp[1], temp[0], tIL['Confidence']], axis=1)
        tIL = tIL.values.tolist()
        tIL_format = [dict(zip(['region', 'position', 'differentiation', 'confidence'], i)) for i in tIL]
        return tIL_format

    def get_FDR_anchor(self,nums):
        nums = sorted(set(nums))
        gaps = [[s, e] for s, e in zip(nums, nums[1:]) if s + 1 < e]
        anchor = [nums[0]] + [i[-1] for i in gaps]
        return anchor