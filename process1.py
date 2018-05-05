#!/usr/bin/env python3

import xml.sax
import sys
from openpyxl import Workbook


def get2nd(v):
    return v[2]

class DumpTagsHandler(xml.sax.ContentHandler):
    row_tags = ['DocumentElement', 'row']
    data_items = [
        ['DocumentElement', 'row', 'ProductName'],
        ['DocumentElement', 'row', 'OrgID'],
        ['DocumentElement', 'row', 'Producer'],
        ['DocumentElement', 'row', 'Place'],
        ['DocumentElement', 'row', 'FarmerName'],
        ['DocumentElement', 'row', 'PackDate'],
        ['DocumentElement', 'row', 'CertificationName'],
        ['DocumentElement', 'row', 'ValidDate'],
        ['DocumentElement', 'row', 'StoreInfo'],
        ['DocumentElement', 'row', 'Tracecode'],
        ['DocumentElement', 'row', 'LandSecNO'],
        ['DocumentElement', 'row', 'ParentTraceCode'],
        ['DocumentElement', 'row', 'TraceCodelist'],
        ['DocumentElement', 'row', 'Log_UpdateTime'],
        ['DocumentElement', 'row', 'OperationDetail'],
        ['DocumentElement', 'row', 'ResumeDetail'],
        ['DocumentElement', 'row', 'ProcessDetail'],
        ['DocumentElement', 'row', 'CertificateDetail']
    ]
    def __init__(self, outwb):
        self.outwb = outwb
        self.parse_state = []
        self.tag_characters = ""

        self.xml_data = [""] * len(self.data_items)

        header = list(map(get2nd, self.data_items))
        self.outwb.active.append(header)

    def startElement(self, name, attrs):
        # print("start: " + name)
        state = name
        self.parse_state.append(state)
        self.tag_characters = ""
        # print("START:", self.parse_state)

    def characters(self, contents):
        self.tag_characters += contents

    def endElement(self, name):
        # print("END:", self.parse_state, self.tag_characters)

        idx = -1
        if self.parse_state in self.data_items:
            idx = self.data_items.index(self.parse_state)
            self.xml_data[idx] = self.tag_characters
        elif self.parse_state == self.row_tags:
            # output data
            # print(self.xml_data)
            self.outwb.active.append(self.xml_data)
            # reset data
            self.xml_data = [""] * len(self.data_items)

        self.parse_state.pop()
        self.tag_characters = ""

# if len(sys.argv) == 2:
if True:
    parser = xml.sax.make_parser()

    outwb = Workbook()

    handler = DumpTagsHandler(outwb)

    parser.setContentHandler(handler)
    parser.parse(open("ResumeData_Plus.xml", "r",  encoding='utf-8'))
    outwb.save("export.xlsx")
else:
    print("add file name as parameter")
