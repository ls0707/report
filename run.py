# -*- coding:utf-8 -*-

from config import Config
from docx_processor import DocxProcessor

cfg=Config('config.ini')
p=DocxProcessor(cfg, 'e:/1/1.docx')
p.start()
