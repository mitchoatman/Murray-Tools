# -*- coding: utf-8 -*-
from pyrevit import script

CONFIG_SECTION = "mypane"
OPT_VISIBLE = "visible"


def get_config():
    return script.get_config(CONFIG_SECTION)


def is_visible():
    cfg = get_config()
    return bool(cfg.get_option(OPT_VISIBLE, False))


def set_visible(value):
    cfg = get_config()
    cfg.set_option(OPT_VISIBLE, bool(value))
    script.save_config()