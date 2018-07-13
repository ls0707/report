# -*- coding:utf-8 -*-

from config import Config
from docx_processor import DocxProcessor

cfg=Config('config.ini')
p=DocxProcessor(cfg, 'd:/1/3.docx')
p.start()
