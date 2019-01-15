#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse

import os
import shutil
import datetime

import pystache
from openpyxl import load_workbook

# data model
class Attribute(object):
    def __init__(self, name: str):
        self._name = name

    def property(self):
        return self._name


class ClassMember(object):
    def __init__(self, attributes: Attribute, data_type: str, name: str):
        self._attributes = attributes
        self._type = data_type
        self._name = name

    def type_name(self):
        return self._type

    def var_name(self):
        return self._name

    def attributes(self):
        return self._attributes


class ClassDeclare(object):
    def __init__(self, name, date, lst):
        self._name = name
        self._classMembers = lst
        self._date = date

    def name(self):
        return self._name

    def class_members(self):
        return self._classMembers

    def date(self):
        return self._date


class ClassList(object):
    def __init__(self, date):
        self._class_list = []
        self._date = date

    def add(self, class_declare):
        self._class_list.append(class_declare)

    def class_list(self):
        return self._class_list

    def date(self):
        return self._date


def set_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", help="path of excel file", required=True)
    parser.add_argument("-t", "--template", help="path of mustache file", required=False)
    parser.add_argument("-o", "--output", help="path of output file", required=True)
    parser.add_argument("-c", "--clean", help="clean up output", action="store_true")
    args = parser.parse_args()
    if not args.input:
        print("PARSER ERROR] arguments is wrong!")
        exit(1)

    return args


def get_class_declare(class_name, primary_index, types, names):
    lst = []
    for i in range(0, len(types)):
        attributes = []
        if i == primary_index:
            attributes.append(Attribute("PrimaryKey"))
        item = ClassMember(attributes, types[i], names[i])
        if item is not None:
            lst.append(item)

    ret = ClassDeclare(class_name, datetime.datetime.now(), lst)
    return ret


args = set_parser()
wb = load_workbook(filename=args.input, data_only=True)

# load mustache template
class_template = ""
class_template_path = "class.mustache"
if args.template:
    class_template_path = args.template

with open(class_template_path) as class_file:
    class_template = class_file.read()

if class_template == "":
    print("class template is empty!")
    exit(1)

output_path = "output"
if args.output:
    output_path = args.output

if os.path.exists(output_path):
    if args.clean:
        shutil.rmtree(output_path, ignore_errors=True)
        os.mkdir(output_path)
else:
    os.mkdir(output_path)

class_list = ClassList(datetime.datetime.now())

for sheet in wb:
    if sheet.title[0] == '_':
        print("INFO] skipped sheet - " + sheet.title)
        continue

    json_dict = {}
    print("INFO] start convert sheet - {}".format(sheet.title))

    data_list = []
    key_list = []

    type_list = []
    val_list = []

    row_index = -1
    primary_index = 0

    for row in sheet:
        row_index = row_index + 1

        # NOTE(jjo): Primary Key를 찾기위한 여정
        if row_index == 2:
            for col in row:
                if col.value is None or col.value == 'PrimaryKey':
                    break
                primary_index = primary_index + 1

        # NOTE(jjo): 자료형을 싹 조사
        elif row_index == 3:
            for col in row:
                if col.value is None:
                    break
                type_list.append(col.value)

        # NOTE(jjo): 변수명을 싹 조사
        elif row_index == 4:
            for col in row:
                if col.value is None:
                    break
                val_list.append(col.value)
        
        # NOTE(jjo): 이후에는 볼 일이 없으므로 끝.
        elif row_index > 4:
            break

    # store class data
    context = get_class_declare(sheet.title, primary_index, type_list, val_list)
    class_list.add(context)

# generate cs file
render_result = pystache.render(class_template, context)

with open(output_path, "w") as class_render:
    class_render.write(render_result)

print("INFO] end convert process - {}".format(sheet.title))

# NOTE(jjo): for test
#with open(sheet.title + '.bson', 'rb') as bson_f:
#    data = bson_f.read()
#    print(bson.loads(data))

