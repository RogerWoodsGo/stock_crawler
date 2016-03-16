#!/usr/bin/python
#coding=utf-8
import urllib2
import urllib
import xlrd
import xlwt
import time
from lxml import etree
from bs4 import BeautifulSoup


def excel_parse(fn):
    person_info = {}
    #global person_info
    excel = xlrd.open_workbook(fn);
    table = excel.sheets()[0]
    print table.nrows, table.ncols
    for row_i  in range(1, table.nrows):
        info = ""
        if(table.cell(row_i, 1).ctype == 2):
            id = str(table.cell(row_i, 1).value).encode("utf-8")
        else:
            id = table.cell(row_i, 1).value.encode("utf-8")
        info += table.cell(row_i, 2).value.encode("utf-8")
        info += ","
        info += table.cell(row_i, 3).value.encode("utf-8")
        info += ","
        info += table.cell(row_i, 4).value.encode("utf-8")
        info += ","
        info += table.cell(row_i, 5).value.encode("utf-8")
        info += ","
        info += table.cell(row_i, 6).value.encode("utf-8")
        info += ","
        info += table.cell(row_i, 7).value.encode("utf-8")
        info += ","
        info += table.cell(row_i, 8).value.encode("utf-8")
        info += ","
        info += str(table.cell(row_i, 9).value).encode("utf-8")
        info += ","
        info += table.cell(row_i, 10).value.encode("utf-8")
        info += ","
        if(table.cell(row_i, 11).ctype == 2):
            info += str(table.cell(row_i, 11).value).encode("utf-8")
        else:
            info += table.cell(row_i, 11).value.encode("utf-8")
        info += ","
        person_info[id] = info
    return person_info

def request_url(url):
    BUFFER_SIZE = 100000
    try:
        resp = urllib2.urlopen(url)
        return resp.read(BUFFER_SIZE)
    except Exception as err:
        return ""

def get_gupiao_list(content, name):
    #soup = BeautifulSoup(content)
    res_list = []
    try:
        tree = etree.HTML(content)
        nodes = tree.xpath(u"//*[@id='table1']")
    #table_title =  nodes[0].xpath(u"tr[2]/td/table[2]/tr[1]/child::*")
        print nodes
        table_content = nodes[0].xpath(u"tr[2]/td/table[2]/tr")
    except Exception as err:
        print err
        table_content = []
        #pass
    for tr in table_content:
        tds = tr.xpath("child::*")
        col = 0
        #print tds[2].xpath("text()")[0]
        #print tds[8].xpath("text()")[0]
        try:
            if tds[8].xpath("text()")[0] == name:
                print tds[2].xpath("text()")[0]
                res_list.append(tds[2].xpath("text()")[0].encode("utf-8"))
                #print tds[8].xpath("text()")[0]
        except Exception as err:
            continue
        #for td in tds:
        #    if(col == 2):
        #        try:
        #            res_list.append(td.xpath("text()")[0])
        #            td.xpath("text()")[0]
        #        except Exception as err:
        #            print "error happened here:", err
        #            continue
        #    col += 1
    return res_list
    #print soup.find_all(id="table1")[0].string.tbody

def excel_write(fn, content):
    with open(fn, "w") as wf:
        for id in content.keys():
            wf.write(id + ",")
            wf.write(content[id] + "\n")

if __name__ == "__main__":
    input_excel_fn = "cj.xls"
    output_excel_fn = "abc.csv"

    person_info = excel_parse(input_excel_fn)
    #query = {"gdmc": u'陈洁'.encode("gbk")}

    #person_info ={"01":""}
    query = {}
    for id in person_info.keys():
        name_with_space = person_info[id].split(",")[0]
        name = "".join(name_with_space.strip(" ").split(" "))
        print name, type(name)
        query["gdmc"] = name.decode('utf-8').encode("gbk")
        print query["gdmc"]
        print type(query["gdmc"])
        url = "http://cwzx.shdjt.com/cwcx.asp?" + urllib.urlencode(query)
        print url
        content = request_url(url).decode("gbk")
        gupiao_list = get_gupiao_list(content, name.decode("utf-8"))
        #print gupiao_list
        person_info[id] = person_info[id] + ";".join(gupiao_list)
        time.sleep(1)
    excel_write(output_excel_fn, person_info)
    print len(content)
