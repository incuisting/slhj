import shutil
from pprint import pprint
import re
from docx import Document
from docx.shared import RGBColor
from os import walk

themes_dic = ['就业', '互联网', '经济', '廉政', '法治', '农业', '传统文化', '精神文化', '安全', '健康', ' 教育', '食品', '科技', '老年人', '留守儿童', '医疗',
              '生态', '创新']


def handle_docx(path):
    doc = Document(path)
    good_words = []
    for i in doc.paragraphs:
        for r in i.runs:
            if r.font.color.rgb != RGBColor(255, 0, 0):
                if '好句' not in r.text:
                    good_words.append(r.text)
    add_words_to_model_file(good_words)


def handle_file_is_done():
    files = []
    for (dirpath, dirnames, filenames) in walk(r'./stuff'):
        files.extend(filenames)
    for f in files:
        if "已改" in f:
            handle_docx(r'./stuff/' + f)


def add_words_to_model_file(words_and_themes):
    current_theme = ''
    for w in words_and_themes:
        matched = re.match(u"[\u4e00-\u9fa5]+", w)
        if matched:
            if matched[0] in themes_dic:
                current_theme = matched[0]
                print(current_theme)


handle_file_is_done()
