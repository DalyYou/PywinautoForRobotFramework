'''
# @Time     : 4/24/2018 3:26 PM
# @Author   : Daly.You
# @Email    : daly.you@wesoft.com
# @File     : __init__.py
# @Project  : WindowsUpdatesLib
# @Python   : 2.7.14
# @Software : PyCharm Community Edition
'''
#!/usr/bin/python
# -*- coding:utf-8 -*-

from WindowsUpdatesHelper import *
version = '1.0'
class WindowsUpdatesLib(WindowsUpdatesHelper,WindowsUpdatesUtilities):
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'