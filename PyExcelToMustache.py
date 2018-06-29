#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse

import os
import shutil
import datetime

import pystache
from xlrd import open_workbook

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


def set_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", help="path of excel file", required=True)
    parser.add_argument("-o", "--output", help="path of output directory", action="store_const", required=False)
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
wb = open_workbook(args.input)

# load mustache template
class_template = ""
with open("class.mustache") as class_file:
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

for sheet in wb.sheets():
    if sheet.name[0] == '_':
        print("INFO] skipped sheet - " + sheet.name)
        continue

    json_dict = {}
    print("INFO] start convert sheet - {}".format(sheet.name))

    data_list = []
    key_list = []

    type_list = []
    val_list = []

    # NOTE(jjo): Primary Key를 찾기위한 여정
    row = sheet.row(2)
    primary_index = 0
    for col in row:
        if col.value == 'PrimaryKey':
            break
        
        primary_index = primary_index + 1

    row = sheet.row(3)
    type_list.append(col.value for col in row)

    row = sheet.row(4)
    val_list.append(col.value for col in row)

    # generate cs file
    context = get_class_declare(sheet.name, primary_index, type_list, val_list)
    render_result = pystache.render(class_template, context)
    with open("{}/class_{}.cs".format(output_path, sheet.name), "w") as class_render:
        class_render.write(render_result)

    print("INFO] end convert process - {}".format(sheet.name))

    # NOTE(jjo): for test
    #with open(sheet.name + '.bson', 'rb') as bson_f:
    #    data = bson_f.read()
    #    print(bson.loads(data))

