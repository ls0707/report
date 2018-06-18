# -*- coding:utf-8 -*-

from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.document import Document as _Document
from docx.table import Table, _Row, _Cell
from docx.text.paragraph import Paragraph
from config import Config
from docx.opc.exceptions import PackageNotFoundError
from clint.textui import prompt, puts, colored, validators
import re


class DocxProcessor(object):
    # ----------------------------- Base Task
    class Task(object):
        def __init__(self):
            self.results = list()

        def _append_result(self):
            pass

    # ----------------------------- Replace Task
    class ParagraphReplaceTask(Task):
        def __init__(self, config):
            self.cfg = config
            super(DocxProcessor.ParagraphReplaceTask, self).__init__()

        def __call__(self, paragraph):
            for old, new in self.cfg.replace_list:
                if old in paragraph.text:
                    # for r in paragraph.runs:
                    #     if old in r.text:
                    paragraph.text = paragraph.text.replace(old, new)
            # print(paragraph.text)

    # ----------------------------- Chapter 4Task
    class ParagraphChapter4Task(Task):
        def __init__(self, config, doc_position):
            self.cfg = config
            self.data = list()
            self.position=doc_position
            super(DocxProcessor.ParagraphChapter4Task, self).__init__()

        def __call__(self, paragraph):
            if self.position.is_chapter4:
                print(self.position.chapter)

        def load_table(self, table):
            self.data.clear()

    # ----------------------------- Cant split Task
    class TableCantsplitTask(Task):
        def __init__(self):
            super(DocxProcessor.TableCantsplitTask, self).__init__()

        def __call__(self, table):
            for row in table.rows:
                print(row)
                tp = parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w')))
                if row._tr.trPr is None:  # trPr为None的情况直接添加属性
                    print(tp, '--added')
                    row._tr.get_or_add_trPr().append(tp)
                else:  # trPr不为None的情况下才做属性遍历，判断该属性是否已经存在
                    for p in row._tr.trPr.getchildren():
                        print(p, '--discovered')
                        if 'cantSplit' not in p.tag:
                            row._tr.get_or_add_trPr().append(tp)

    # ----------------------------- Chapter6 Task
    class TableChapter6Task(Task):
        def __init__(self):
            super(DocxProcessor.TableChapter6Task, self).__init__()

        def __call__(self, table):
            pass

    # ----------------------------- Position
    class DocPosition(object):
        def __init__(self):
            self.__ro_chapter = re.compile('\d+(\.\d+)+')
            self.__chapter = list()
            self.this_chapter = list()

        def location_from_paragraph(self, paragraph):
            ma_chapter = self.__ro_chapter.match(paragraph.text)

            if ma_chapter:
                self.__chapter = ma_chapter.group().split('.')

        @property
        def chapter(self):
            return self.__chapter

        def is_chapter4(self):
            result = False
            if len(self.__chapter) > 0:
                if self.__chapter[0] == 4:
                    result = True
            return result

    # ============================= DocxProcessor
    def __init__(self, processor_cfg, docx_file):
        assert isinstance(processor_cfg, Config)
        self.file = docx_file
        self.cfg = processor_cfg
        self.document = None
        try:
            self.document = Document(docx_file)
        except PackageNotFoundError as e:
            print(e)
            exit(0)
        self.position=DocxProcessor.DocPosition()
        self.paragraph_replace_task = DocxProcessor.ParagraphReplaceTask(processor_cfg)
        self.paragraph_chapter4_task = DocxProcessor.ParagraphChapter4Task(processor_cfg, self.position)
        self.table_cantsplit_task = DocxProcessor.TableCantsplitTask()
        self.table_chapter6_task = DocxProcessor.TableChapter6Task()

        self.paragraph_process_list = [self.paragraph_replace_task if self.cfg.replace_enabled else None,
                                       self.paragraph_chapter4_task if self.cfg.chapter4_enabled else None
                                       ]
        self.table_process_list = [self.table_cantsplit_task if self.cfg.table_cant_split else None,
                                   self.table_replace_func if self.cfg.replace_enabled else None,
                                   self.paragraph_chapter4_task.load_table if self.cfg.chapter4_enabled else None,
                                   self.table_chapter6_task if self.cfg.chapter6_enabled else None
                                   ]

    def start(self):  # 执行文档处理过程，处理结束后保存
        for block in self.iter_block_items(self.document):
            if isinstance(block, Paragraph):
                self.process_paragraph(block)

            elif isinstance(block, Table):
                self.process_table(block)

        inst_options = [{'selector': '1', 'prompt': '确认保存', 'return': False},
                        {'selector': '2', 'prompt': '退出/不保存', 'return': True}]
        puts(colored.yellow('=' * 50))
        quit = prompt.options('以上内容将发生改变，请确认：', inst_options)
        if quit:
            puts(colored.green('变更未作保存。'))
            exit(0)
        while True:
            try:
                self.document.save(self.file)  # 保存处理完毕的文件
                puts(colored.red('变更已保存。'))
                exit(0)
            except PermissionError as e:
                print(e)
                inst_options = [{'selector': '1', 'prompt': '重试', 'return': False},
                                {'selector': '2', 'prompt': '退出/不保存', 'return': True}]
                quit = prompt.options('文件是否已经打开？请关闭后重试。', inst_options)
                if quit:
                    exit(0)

    def process_paragraph(self, paragraph):
        self.position.location_from_paragraph(paragraph)
        for func in self.paragraph_process_list:  # 顺次以paragraph为参数，执行列表中的函数
            if func is not None:
                func(paragraph)
        #
        # if 'a' in paragraph.text:
        #     print('in')
        #     for i in paragraph.runs:
        #         if 'a' in i.text:
        #             i.text = i.text.replace('a', 'd')

    def process_table(self, table):
        for func in self.table_process_list:
            if func is not None:
                func(table)

    def table_replace_func(self, table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    self.paragraph_replace_task(paragraph)

    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """

    def iter_block_items(self, parent):
        if isinstance(parent, _Document):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        elif isinstance(parent, _Row):
            parent_elm = parent._tr
        else:
            raise ValueError("something's not right")
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)
