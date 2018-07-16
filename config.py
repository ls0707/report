# -*- coding:utf-8 -*-

from configparser import ConfigParser, NoOptionError, ParsingError


class Config(object):
    def __init__(self, configfile):
        self.__chapter4 = False
        self.__chapter6 = False
        self.__replace = False
        self.__chapter4_mode_int = 0
        self.__replace_str = list()
        self.__cant_split = True
        config = ConfigParser()
        try:
            config.read(configfile)
            self.__chapter4 = config.getboolean('options', 'chapter4')
            self.__chapter6 = config.getboolean('options', 'chapter6')
            self.__replace = config.getboolean('options', 'replace')
            self.__chapter4_mode_int = config.getint('chapter4', 'mode')
            self.__cant_split = config.getboolean('options', 'table_cantsplit')
            for opt in config.options('replace'):
                self.__replace_str.append(config.get('replace', opt).split('->', 1))

        except (ParsingError, NoOptionError, ValueError) as e:
            print(e)
            exit(-1)

    @property
    def chapter4_enabled(self):
        return self.__chapter4

    @property
    def chapter6_enabled(self):
        return self.__chapter6

    @property
    def chapter4_mode(self):
        return self.__chapter4_mode_int

    @property
    def replace_enabled(self):
        return self.__replace

    @property
    def table_cant_split(self):
        return self.__cant_split

    @property
    def replace_list(self):
        return self.__replace_str
