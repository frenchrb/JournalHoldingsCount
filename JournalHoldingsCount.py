import sys
import configparser
import xlrd
import xlwt
import xlutils.copy
import requests
import lxml
from lxml import etree

#input is path to Excel spreadsheet
def main(input):
    #read config file
    config = configparser.ConfigParser()
    config.read('local_settings.ini')
    wskey = config['WorldCat Search API']['wskey']
    
    #read spreadsheet
    book_in = xlrd.open_workbook(input)
    sheet1 = book_in.sheet_by_index(0) #get first sheet
    sheet1_name = book_in.sheet_names()[0] #name of first sheet
    #print('sheet1 type:',type(sheet1))
    sheet1_col_headers = sheet1.row_values(0)
    #print(sheet1_col_headers)
    
    #Column U = list of holding libraries = sheet1.row_values(0)[20]
    #print(sheet1.row_values(0)[20])
    #Column V = count of holding libraries = sheet1.row_values(0)[21]
    #print(sheet1.row_values(0)[21])
    
    #turn the xlrd Book into xlwt Workbook
    book_out = xlutils.copy.copy(book_in)
    
    #add new column headers
    book_out.get_sheet(0).write(0,20,'Holding Libraries')
    book_out.get_sheet(0).write(0,21,'Holding Libraries Count')
    
    item_col_index = 0
    eISSN_col_index = 7
    ISSN_col_index = 8
    
    #for row in range(1, sheet1.nrows):
    for row in range(1, sheet1.nrows):
        print('Item', sheet1.cell(row, item_col_index).value)
        eISSN = sheet1.cell(row,eISSN_col_index).value
        ISSN = sheet1.cell(row,ISSN_col_index).value
        #print('eISSN:', eISSN)
        #print('ISSN:', ISSN)
        
        if eISSN or ISSN:
            oclc_num_list = []
            try:
                #search eISSN to get OCLC numbers
                if eISSN:
                    response = requests.get('http://www.worldcat.org/webservices/catalog/search/sru?query=srw.in+all+'+eISSN+'&servicelevel=full&maximumRecords=100&wskey='+wskey+'&recordSchema=info%3Asrw%2Fschema%2F1%2Fmarcxml&frbrGrouping=off')
                    outfile = open ('results.txt', 'w', encoding='utf-8')
                    if response.status_code == 200:
                        #no problems writing to a file correctly, but can't print to console if not utf-8 and prints incorrectly if utf-8
                        #print(response.text.encode('utf-8'))
                        outfile.write(response.text)
                        
                        #parse xml
                        tree = etree.parse('results.txt')
                        #records is a list of tree elements
                        records = tree.xpath('/srw:searchRetrieveResponse/srw:records/srw:record/srw:recordData/marc:record', namespaces={'srw': 'http://www.loc.gov/zing/srw/', 'marc': 'http://www.loc.gov/MARC21/slim'})
                        #print(records)
                        
                        for r in records:
                            m040b = r.xpath('marc:datafield[@tag="040"]/marc:subfield[@code="b"]/text()', namespaces={'marc': 'http://www.loc.gov/MARC21/slim'})
                            #print(m040b)
                            
                            #if 040$b is eng or doesn't exist, save 001
                            #TODO might also need to capture records where 040$b is blank/doesn't exist
                            if 'eng' in m040b:
                                m001 = r.xpath('marc:controlfield[@tag="001"]/text()', namespaces={'marc': 'http://www.loc.gov/MARC21/slim'})
                                oclc_num_list.append(m001[0])     
                
                #search ISSN to get OCLC numbers
                if ISSN:
                    response = requests.get('http://www.worldcat.org/webservices/catalog/search/sru?query=srw.in+all+'+ISSN+'&servicelevel=full&maximumRecords=100&wskey='+wskey+'&recordSchema=info%3Asrw%2Fschema%2F1%2Fmarcxml&frbrGrouping=off')
                    outfile = open ('results.txt', 'w', encoding='utf-8')
                    if response.status_code == 200:
                        outfile.write(response.text)
                        tree = etree.parse('results.txt')
                        records = tree.xpath('/srw:searchRetrieveResponse/srw:records/srw:record/srw:recordData/marc:record', namespaces={'srw': 'http://www.loc.gov/zing/srw/', 'marc': 'http://www.loc.gov/MARC21/slim'})
                        
                        for r in records:
                            m040b = r.xpath('marc:datafield[@tag="040"]/marc:subfield[@code="b"]/text()', namespaces={'marc': 'http://www.loc.gov/MARC21/slim'})
                            #print(m040b)
                            
                            #if 040$b is eng or doesn't exist, save 001
                            #TODO might also need to capture records where 040$b is blank/doesn't exist
                            if 'eng' in m040b:
                                m001 = r.xpath('marc:controlfield[@tag="001"]/text()', namespaces={'marc': 'http://www.loc.gov/MARC21/slim'})
                                oclc_num_list.append(m001[0])
                print('OCLC numbers:', oclc_num_list)
                print('Getting holdings...')
                
                #for each OCLC number, get holdings and concat, put into spreadsheet col.U
                #count holdings, put into spreadsheet col.V
                libraries_list = []
                libraries_count = 0
                
                for o in oclc_num_list:
                    response = requests.get('http://www.worldcat.org/webservices/catalog/content/libraries/'+o+'?location=virginia&libtype=1&wskey='+wskey+'&format=json&maximumLibraries=100&servicelevel=full')
                    if response.status_code == 200:
                        #print(response.text)
                        json = response.json()
                        #print(json['library'])
                        
                        if 'library' in json:
                            for l in json['library']:
                                if 'institutionName' in l:
                                    #print(l['institutionName'])
                                
                                    #add library to list if not already in list
                                    if l['institutionName'] not in libraries_list:
                                        libraries_count += 1
                                        libraries_list.append(l['institutionName'])
            except etree.XMLSyntaxError:
                print('XMLSyntaxError; check manually')
                libraries_list = ['check manually']
                libraries_count = 'check manually'
            except ValueError:
                print('ValueError; check manually')
                libraries_list = ['check manually']
                libraries_count = 'check manually'
            
            book_out.get_sheet(0).write(row,20,','.join(libraries_list))
            book_out.get_sheet(0).write(row,21,libraries_count)
        
            print('Holdings count:', libraries_count)
        #print(libraries_list)
        book_out.save(input+'_new.xls')
        print('------------------------------')
    
if __name__ == '__main__':
    main(sys.argv[1])
