import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from zipfile import BadZipFile
import olefile
import csv
import networkx as nx
import matplotlib.pyplot as plt
from datetime import datetime

search_file_name_list = []
search_file_path_list = []
finalFileList = []

class FileInformaiton:
    def __init__(self):
        self.filename = ''
        self.filepath = ''
        self.company = ''
        self.author = ''
        self.lastsaved = ''
        self.layout = LayoutList()
        self.rsid = []
        self.relatedFile = []
        self.duplicated = False
        self.modifiedtime = ''
        self.modifiedtime_filesystem = 0

class LayoutList:
    def __init__(self):
        self.pgSz_w = 0
        self.pgSz_h = 0
        self.pgMar_g = 0
        self.pgMar_f = 0
        self.pgMar_h = 0
        self.pgMar_l = 0
        self.pgMar_b = 0
        self.pgMar_r = 0
        self.pgMar_t = 0

def Searchfile(dirname):
    try:
        filenames = os.listdir(dirname)
        for filename in filenames:
            full_filename = os.path.join(dirname, filename)
            if os.path.isdir(full_filename):
                Searchfile(full_filename)
            else:
                ext = os.path.splitext(full_filename)[-1]
                if (ext == '.docx' or ext == '.doc' or ext == '.xlsx' or ext == '.xls' or ext == '.ppt' or ext == '.pptx') and '$' not in full_filename:
                    search_file_name_list.append(filename)
                    search_file_path_list.append(full_filename)

    except PermissionError:
        pass

    return 0

def ExtractInformation(search_file_name_list, search_file_path_list, path):
    file_list = []
    for i in range(0, len(search_file_path_list)):
        file_info = FileInformaiton()
        file_list.append(file_info)
        file_list[i].filename = search_file_name_list[i]
        file_list[i].filepath = search_file_path_list[i]
        file_list[i].modifiedtime_filesystem = os.stat(search_file_path_list[i]).st_mtime_ns

    for file_list_single in file_list:
        ext = os.path.splitext(file_list_single.filename)[-1]
        # OOXML
        if ext == '.docx' or ext == '.xlsx' or ext == '.pptx':
            try:
                with ZipFile(file_list_single.filepath) as zf:
                    fantasy_zip = zipfile.ZipFile(file_list_single.filepath)
                    fantasy_zip.extractall(output_path_dir + '/' + now + '_Result/extracted')
                    fantasy_zip.close()

                    # docx
                    if ext == '.docx':
                        # Extracting RSID in document.xml
                        tree = ET.parse(output_path_dir + '/' + now + '_Result/extracted/word/document.xml')
                        root = tree.getroot()
                        product = root[0]

                        for vendor in product:
                            if vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR") != None:
                                file_list_single.rsid.append(
                                    vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR"))
                            if vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault") != None:
                                file_list_single.rsid.append(
                                    vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault"))
                            if vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr") != None:
                                file_list_single.rsid.append(
                                    vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr"))
                            if vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidP") != None:
                                file_list_single.rsid.append(
                                    vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidP"))
                        for second_vendor in vendor:
                            if second_vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR") != None:
                                file_list_single.rsid.append(
                                    second_vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR"))
                            if second_vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault") != None:
                                file_list_single.rsid.append(
                                    second_vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault"))
                            if second_vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr") != None:
                                file_list_single.rsid.append(
                                    second_vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr"))
                            if second_vendor.attrib.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidP") != None:
                                file_list_single.rsid.append(
                                    second_vendor.attrib.get(
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidP"))

                        file_list_single.rsid = list(set(file_list_single.rsid))  # Removing duplicated RSID

                        # extracting layout in document.xml
                        for vendor in product:
                            if vendor.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr":
                                for layout_vendor in vendor:
                                    if layout_vendor.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgSz":
                                        file_list_single.layout.pgSz_w = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w")
                                        file_list_single.layout.pgSz_h = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}h")
                                    if layout_vendor.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgMar":
                                        file_list_single.layout.pgMar_t = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top")
                                        file_list_single.layout.pgMar_r = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right")
                                        file_list_single.layout.pgMar_b = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom")
                                        file_list_single.layout.pgMar_l = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left")
                                        file_list_single.layout.pgMar_h = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}header")
                                        file_list_single.layout.pgMar_f = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footer")
                                        file_list_single.layout.pgMar_g = layout_vendor.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gutter")

                    # app.xml (Company)
                    tree = ET.parse(output_path_dir + '/' + now + '_Result/extracted/docProps/app.xml')
                    root = tree.getroot()

                    for vendor in root:
                        if vendor.tag == "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company":
                            location = vendor.tag.find('}')
                            metadata_type = vendor.tag[location + 1:]
                            if (vendor.text == None):
                                file_list_single.company = "None"
                            else:
                                file_list_single.company = vendor.text

                    # core.xml (creator, lastmodifiedby)
                    tree = ET.parse(output_path_dir + '/' + now + '_Result/extracted/docProps/core.xml')
                    root = tree.getroot()

                    for vendor in root:
                        if vendor.tag == "{http://purl.org/dc/elements/1.1/}creator":
                            location = vendor.tag.find('}')
                            metadata_type = vendor.tag[location + 1:]
                            if (vendor.text == None):
                                file_list_single.author = "None"
                            else:
                                file_list_single.author = vendor.text
                        if vendor.tag == "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy":
                            location = vendor.tag.find('}')
                            metadata_type = vendor.tag[location + 1:]
                            if (vendor.text == None):
                                file_list_single.lastsaved = "None"
                            else:
                                file_list_single.lastsaved = vendor.text
                        if vendor.tag == "{http://purl.org/dc/terms/}modified":
                            location = vendor.tag.find('}')
                            metadata_type = vendor.tag[location + 1:]
                            if (vendor.text == None):
                                file_list_single.modifiedtime = "None"
                            else:
                                file_list_single.modifiedtime = vendor.text

            except BadZipFile:
                print('———————————————Damaged document——————————————————')
                print(file_list_single.filepath)
                print('—————————————————————————————————————————————————')
                continue
        #doc, xls, ppt
        elif ext == '.doc' or ext == '.xls' or ext == '.ppt':
            of = olefile.OleFileIO(file_list_single.filepath)
            meta = of.get_metadata()
            if meta.author != None:
                file_list_single.author = (meta.author).decode('cp949')
            elif meta.last_saved_by != None:
                file_list_single.lastsaved = meta.last_saved_by.decode('cp949')
            elif meta.company != None:
                file_list_single.company = meta.company.decode('cp949')
        else:
            print('—————————————Not a MS Office file————————————————')
            print(file_list_single.filepath)
            print('—————————————————————————————————————————————————')

    shutil.rmtree(output_path_dir + '/' + now + '_Result/extracted/')
    return file_list

def compareRSID(first, second):
    for a in first:
        if a in second:
            return 1
        else:
            pass
    return 0

def sameRSID(first, second):
    for a in first:
        if a in second:
            return a

def layoutCompare(a,b):
    if a.layout.pgMar_h == b.layout.pgMar_h and a.layout.pgMar_l == b.layout.pgMar_l and a.layout.pgMar_r == b.layout.pgMar_r and a.layout.pgMar_t == b.layout.pgMar_t and a.layout.pgMar_b == b.layout.pgMar_b and a.layout.pgMar_f == b.layout.pgMar_f and a.layout.pgMar_g == b.layout.pgMar_g and a.layout.pgSz_h == b.layout.pgSz_h and a.layout.pgSz_w == b.layout.pgSz_w:
        if a.layout.pgMar_l == '1440' and a.layout.pgMar_r == '1440' and a.layout.pgMar_t == '1701' and a.layout.pgMar_b == '851' and a.layout.pgMar_f == '992' and a.layout.pgMar_g == '0' and a.layout.pgSz_h == '16838' and a.layout.pgSz_w == '11906':
            return 0
        elif a.layout.pgMar_l == '1440' and a.layout.pgMar_r == '1440' and a.layout.pgMar_t == '1440' and a.layout.pgMar_b == '1440' and a.layout.pgMar_f == '720' and a.layout.pgMar_h == '720' and a.layout.pgMar_g == '0' and a.layout.pgSz_h == '15840' and a.layout.pgSz_w == '12240':
            return 0
        elif a.layout.pgMar_l == '1800' and a.layout.pgMar_r == '1800' and a.layout.pgMar_t == '1440' and a.layout.pgMar_b == '1440' and a.layout.pgMar_f == '720' and a.layout.pgMar_h == '720' and a.layout.pgMar_g == '0' and a.layout.pgSz_h == '15840' and a.layout.pgSz_w == '12240':
            return 0
        elif a.layout.pgMar_l == '720' and a.layout.pgMar_r == '720' and a.layout.pgMar_t == '720' and a.layout.pgMar_b == '720' and a.layout.pgMar_f == '720' and a.layout.pgMar_h == '720' and a.layout.pgMar_g == '0' and a.layout.pgSz_h == '15840' and a.layout.pgSz_w == '12240':
            return 0
        elif a.layout.pgMar_l == '1440' and a.layout.pgMar_r == '1440' and a.layout.pgMar_t == '1440' and a.layout.pgMar_b == '1440' and a.layout.pgMar_f == '1440' and a.layout.pgMar_h == '1440' and a.layout.pgMar_g == '0' and a.layout.pgSz_h == '15840' and a.layout.pgSz_w == '12240':
            return 0
        elif a.layout.pgMar_l == 0 and a.layout.pgMar_r == 0 and a.layout.pgMar_t == 0 and a.layout.pgMar_b == 0 and a.layout.pgMar_f == 0 and a.layout.pgMar_h == 0 and a.layout.pgMar_g == 0 and a.layout.pgSz_h == 0 and a.layout.pgSz_w == 0:
            return 0
        else:
            return 1
    else:
        return 0

def insertionSort(x):
    for i in range(1, len(x)):
        if str(x[i-1].modifiedtime) != '':
            while i > 0 and datetime.strptime(x[i-1].modifiedtime, "%Y-%m-%dT%H:%M:%SZ") > datetime.strptime(x[i].modifiedtime, "%Y-%m-%dT%H:%M:%SZ"):
                temp = x[i - 1]
                x[i - 1] = x[i]
                x[i] = temp
                i = i - 1
    return x

def Recur_clustering_function(Clustered_group, relatedFile):
    i = 0
    for i in range(0, len(relatedFile.relatedFile)):
        if relatedFile.relatedFile[i] in Clustered_group:
            continue
        else:
            Clustered_group.append(relatedFile.relatedFile[i])
            Recur_clustering_function(Clustered_group, relatedFile.relatedFile[i])

def ClusteringFuction(finalFileList, optionNumber):
    for a in finalFileList:
        for b in finalFileList:
            if a != b:
            # Company
                if optionNumber == '1':
                    if (a.company == b.company and a.company != 'None' and a.company != ' ' and a.company != ''):
                        a.relatedFile.append(b)
            # Author
                elif optionNumber == '2':
                    if (a.author == b.author and a.author != 'None' and a.author != ' ' and a.author != ''):
                        a.relatedFile.append(b)
            # Last Saved By
                elif optionNumber == '3':
                    if (a.lastsaved == b.lastsaved and a.lastsaved != 'None' and a.lastsaved != ' ' and a.lastsaved != ''):
                        a.relatedFile.append(b)
            # Company, Author, Last Saved By
                elif optionNumber == '4':
                    if (a.company == b.company and a.company != 'None' and a.company != ' ' and a.company != '') or (a.author == b.author and a.author != 'None' and a.author != ' ' and a.author != '') or (a.lastsaved == b.lastsaved and a.lastsaved != 'None' and a.lastsaved != ' ' and a.lastsaved != ''):
                        a.relatedFile.append(b)
            # Layout
                elif optionNumber == '5':
                    if layoutCompare(a,b) == 1:
                        a.relatedFile.append(b)
            # RSID
                elif optionNumber == '6':
                    if compareRSID(a.rsid, b.rsid) == 1:
                        a.relatedFile.append(b)
            # Everything
                elif optionNumber == '7':
                    if (a.company == b.company and a.company != 'None' and a.company != ' ' and a.company != '') or (a.author == b.author and a.author != 'None' and a.author != ' ' and a.author != '') or (a.lastsaved == b.lastsaved and a.lastsaved != 'None' and a.lastsaved != ' ' and a.lastsaved != '') or (layoutCompare(a,b) == 1) or (compareRSID(a.rsid, b.rsid) == 1):
                        a.relatedFile.append(b)
    relatedFileList = []
    notrelatedFileList = []

    for temp in finalFileList:
        if len(temp.relatedFile) == 0:
            notrelatedFileList.append(temp)
        else:
            relatedFileList.append(temp)

    Clustered_group = []
    for i in range(0, len(relatedFileList)):
        Clustered_group_file_list = []
        Clustered_group.append(Clustered_group_file_list)
        Clustered_group[i].append(relatedFileList[i])
        Recur_clustering_function(Clustered_group[i], relatedFileList[i])
        #insertionSort(Clustered_group[i])

    Clustered_group_final = []
    for i in Clustered_group:
        duplicated = 0
        if len(Clustered_group_final) > 0:
            for j in Clustered_group_final:
                if set(i) == set(j):
                    duplicated = 0
                    break
                else:
                    duplicated = 1
            if duplicated == 1:
                Clustered_group_final.append(i)
        else:
            Clustered_group_final.append(i)

    print("[Completed grouping documents]")

    with open(output_path_dir + '/' + now + '_Result/result.csv', 'w', newline='') as cf:
        wr = csv.writer(cf)
        for temp in range(0, len(Clustered_group_final)):
            wr.writerow(['Group '+str(temp)])
            if not (os.path.isdir(output_path_dir + '/' + now + '_Result/GROUP'+str(temp))):
                os.makedirs(os.path.join(output_path_dir + '/' + now + '_Result/GROUP'+str(temp)))
            for i in range(0, len(Clustered_group_final[temp])):
                wr.writerow([Clustered_group_final[temp][i].filepath, Clustered_group_final[temp][i].company, Clustered_group_final[temp][i].author, Clustered_group_final[temp][i].lastsaved])
                shutil.copy(Clustered_group_final[temp][i].filepath, output_path_dir + '/' + now + '_Result/GROUP'+str(temp))
            wr.writerow([' '])
    print(" ")
    print("---------------[Result]---------------")
    print("1) Total number of documents: ", len(finalFileList))
    print("2) Number of unrelated documents: ", len(notrelatedFileList))
    print("3) Number of related documents: ", len(relatedFileList))
    print("4) Number of documents groups: ", len(Clustered_group_final))
    print("--------------------------------------")
    return notrelatedFileList, Clustered_group_final, relatedFileList

def Visualization(notrelatedFileList, Clustered_group_final, relatedFileList):
    G = nx.DiGraph()

    # 0 = same
    # 1 = not same

    for file in notrelatedFileList:
        G.add_node(file)

    for cluster in Clustered_group_final:
        for i in cluster:
            for j in i.relatedFile:
                if i.modifiedtime == j.modifiedtime:
                    if i.modifiedtime_filesystem > j.modifiedtime_filesystem:
                        G.add_edge(i.filepath, j.filepath, weight=0)
                    else:
                        G.add_edge(j.filepath, i.filepath, weight=0)
                elif datetime.strptime(i.modifiedtime, "%Y-%m-%dT%H:%M:%SZ") > datetime.strptime(j.modifiedtime, "%Y-%m-%dT%H:%M:%SZ"):
                    G.add_edge(i.filepath, j.filepath, weight=1)
                elif datetime.strptime(i.modifiedtime, "%Y-%m-%dT%H:%M:%SZ") < datetime.strptime(j.modifiedtime, "%Y-%m-%dT%H:%M:%SZ"):
                    G.add_edge(j.filepath, i.filepath, weight=1)

    esame = [(u, v) for (u, v, d) in G.edges(data=True) if d['weight'] == 0]
    ediff = [(u, v) for (u, v, d) in G.edges(data=True) if d['weight'] == 1]
    pos = nx.spring_layout(G)  # positions for all nodes

    nx.draw_networkx_nodes(G, pos, node_color='black', node_size=10)
    nx.draw_networkx_edges(G, pos, edgelist=ediff, width=1.0, edge_color='green', arrowsize=10)
    nx.draw_networkx_edges(G, pos, edgelist=esame, width=1.0, edge_color='red', arrowsize=10)
    plt.axis('off')
    plt.savefig(output_path_dir + '/' + now + '_Result/weighted_graph.png')  # save as
    plt.show()

# main
now = datetime.now()
now = "%s-%s-%s_%s-%s-%s" % (now.year, now.month, now.day, now.hour, now.minute, now.second)

#os.makedirs(os.path.join('./'+now+'_Result'))
path_dir = input("Type directory path : ")
output_path_dir = input("Type output path : ")
if not (os.path.isdir(output_path_dir + '/' + now + '_Result/')):
    os.makedirs(os.path.join(output_path_dir + '/' + now + '_Result/'))

Searchfile(path_dir)
#Searchfile('./sample')

print(" ")
print("[Feature List]")
print("1. Company")
print("2. Author")
print("3. Last Saved By")
print("4. Layout")
print("5. RSID")
print("6. Company, Author, Last Saved By")
print("7. Everything")
print(" ")
optionNumber = input("Select the feature: ")
print(" ")

print("[Extracting documents information...]")
finalFileList = ExtractInformation(search_file_name_list, search_file_path_list, output_path_dir)
notrelatedFileList, finalClutered, relatedFileList = ClusteringFuction(finalFileList, optionNumber)
#notrelatedFileList, finalClutered, relatedFileList = ClusteringFuction(finalFileList, '7')
Visualization(notrelatedFileList, finalClutered, relatedFileList)
