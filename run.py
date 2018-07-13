# -*- coding:utf-8 -*-

from config import Config
from docx_processor import DocxProcessor
import click


def print_version(ctx, param, value):
    if not value or ctx.resilient_parsing:
        return
    click.echo('Version 1.1')
    ctx.exit()


@click.command()
@click.option('--version', is_flag=True, callback=print_version, expose_value=False, is_eager=True)
# @click.option('--output', default='newdatabase.sqlite', help='指定想要输出的DB文件名。')
@click.argument('docx_file')
def process_docx(docx_file):
    cfg = Config('config.ini')
    p = DocxProcessor(cfg, docx_file)
    p.start()


process_docx()
