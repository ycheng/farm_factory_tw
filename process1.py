#!/usr/bin/env python3

import xml.sax
import sys
from openpyxl import Workbook
import pandas

def get0th(v):
    return v[0]

def get1st(v):
    return v[1]

def get2nd(v):
    return v[2]

def toStr(v):
    return str(v)

def loc2locstr(feature):
    try:
        u = feature["geometry"]["coordinates"][0][0]
        u1 = ",".join(map(toStr,u))
        # print(u1)
        p = feature["properties"]
        xmax = p["xmax"]
        ymax = p["ymax"]
        xmin = p["xmin"]
        ymin = p["ymin"]
        xcenter = p["xcenter"]
        ycenter = p["ycenter"]
        p1 = ",".join(map(toStr,[xmax, xmin, ymax, ymin, xcenter, ycenter]))
        # print(p1)
        return [u1, p1]
    except TypeError:
        return None

# TODO: local cache LOG data.
def getLOCs(landSecNos):
    if len(landSecNos) == 0:
        return ["",""]
    landSecNosAr = landSecNos.split(";")
    print(len(landSecNosAr))
    print('=' * 10)
    query_str = "&".join(map(lambda x: "lands[]=" + x, landSecNos.split(";")))
    if len(query_str) == 0:
        return ""
    print(query_str)
    url = "http://twland.ronny.tw/index/search?" + query_str
    print(url)
    json = pandas.read_json(url)
    us = list(filter(lambda x: x is not None, map(loc2locstr, json["features"])))
    ret = [";".join(map(get0th, us)), ";".join(map(get1st, us))] 
    # print(p1)
    return ret

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
        header.append("locs")
        header.append("xmax,xmin,ymax,ymin,xcenter,ycenter")
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
            landSecNO = self.xml_data[10]
            locs = getLOCs(landSecNO)
            # land
            self.outwb.active.append(self.xml_data + locs)
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
    # getLOC("QB0317,01120000")
    print(getLOCs("PA0045,49730000;PA0045,49720010;PA0045,58160000;PD0412,25130000;PD0405,05320000;PA0043,09940000"))
    print("add file name as parameter")
