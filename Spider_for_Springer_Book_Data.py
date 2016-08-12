# -*- coding: utf-8 -*-
"""
Created on Thu Aug 11 18:55:41 2016

@author: WHUER
"""

import urllib2
import re
import datetime
import traceback
import csv
import xlsxwriter
reload(sys)
sys.setdefaultencoding('utf8')

starttime = datetime.datetime.now()
###############################################################################
#                                   子程序                                     #
###############################################################################

def get_data(url):
    req = urllib2.urlopen(url)
    page = req.read()
    req.close()
    return page
	
def options_output(options_list):
    for i in xrange(len(options_list)):
        print options_list[i]

def check_input(input_tip, max_num):
    input_str = '1'
    input_str = raw_input('\nPlease Input The ' + input_tip + ':')
    print ''
    if input_str == '':
        input_str = '1'
    if input_str.isdigit():
        input_str = int(input_str)
        if input_str < 1 or input_str > max_num:
            print 'The ' + input_tip + ' You Input Is Out of Range! Please Try Again!'
            check_input(input_tip, max_num)
    else:
        print 'The ' + input_tip + ' You Input Is Not A Num! Please Try Again!'
        check_input(input_tip, max_num)
    return input_str
###############################################################################
#                                   主程序                                     #
###############################################################################

Discipline_options_list = ['[1] Engineering','[2] Computer Science','[3] Medicine','[4] Social Sciences','[5] Mathematics','[6] Physics','[7] Life Sciences','[8] Biomedical Sciences','[9] Chemistry','[10]Materials','[11]Education & Language','[12]Earth Sciences and Geography','[13]Environment','[14]Popular Science','[15]Psychology','[16]Law','[17]Philosophy','[18]Statistics','[19]Astronomy','[20]Public Health','[21]Energy','[22]Food Science & Nutrition','[23]Climate','[24]Business & Management','[25]Water','[26]Economics','[27]Earth Sciences & Geography','[28]Environmental Sciences','[29]Architecture & Design']
options_output(Discipline_options_list)
input_tip = 'Num of Discipline'
max_num = 29
Discipline_id = check_input(input_tip, max_num)
discipline_list = ['Engineering','Computer+Science','Medicine','Social+Sciences','Mathematics','Physics','Life+Sciences','Biomedical+Sciences','Chemistry','Materials','Education+%26+Language','Earth+Sciences+and+Geography','Environment','Popular+Science','Psychology','Law','Philosophy','Statistics','Astronomy','Public+Health','Energy','Food+Science+%26+Nutrition','Climate','Business+%26+Management','Water','Economics','Earth+Sciences+%26+Geography','Environmental+Sciences','Architecture+%26+Design']
Discipline_list = ['Engineering','Computer Science','Medicine','Social Sciences','Mathematics','Physics','Life Sciences','Biomedical Sciences','Chemistry','Materials','Education & Language','Earth Sciences and Geography','Environment','Popular Science','Psychology','Law','Philosophy','Statistics','Astronomy','Public Health','Energy','Food Science & Nutrition','Climate','Business & Management','Water','Economics','Earth Sciences & Geography','Environmental Sciences','Architecture & Design']
discipline = discipline_list[Discipline_id - 1]
Discipline = Discipline_list[Discipline_id - 1]
Language_options_list = ['[1]English','[2]German','[3]Dutch','[4]French','[5]Italian','[6]Spanish','[7]Portuguese','[8]Chinese']
Language_name_list = ['English','German','Dutch','French','Italian','Spanish','Portuguese','Chinese']
options_output(Language_options_list)
input_tip = 'Num of Language'
max_num = 8
Language_id = check_input(input_tip, max_num)
Language_list = ['En','De','Nl','Fr','It','Es','Pt','Zh']
Language_code = Language_list[Language_id - 1]
Language = Language_name_list[Language_id - 1]

search_page_num = '1'
Content_type = 'Book' 
search_url = 'http://link.springer.com/search/page/' + search_page_num + '?facet-discipline=%22'\
            + discipline + '%22&facet-language=%22' + Language_code + '%22&' + '&facet-content-type=%22'\
            + Content_type + '%22'

search_page = get_data(search_url)
search_page = "".join(search_page.split())
result_num = re.findall('class="number-of-search-results-and-search-terms"><strong>(.*?)</strong>',search_page)
result_num = int(re.sub(',', '', result_num[0]))
last_item = result_num %20
page_num = re.findall('class="number-of-pages">(.*?)</span>',search_page)
page_num = int(re.sub(',', '', page_num[0]))
input_tip = 'Start Num of Result List Page'
max_num = page_num
i_start = check_input(input_tip, max_num) - 1
input_tip = 'Start Num of ' + str(i_start +1) + 'st Result List Page`s Item'
max_num = 20
j_start = check_input(input_tip, max_num) - 1
parameter_name_list = ['Book Title', 'Book Subtitle', 'Copyright', 'Authors', 'DOI', 'Print ISBN', 'Online ISBN', 'Publisher', 'Copyright Holder', 'Discipline', 'book url', 'Language', 'Content Type', 'citation count', 'altmetric mention count', 'reader count', 'review count', 'download count', 'query time', 'result id']
location_list = ['A:A','B:B','C:C','D:D','E:E','F:F','G:G','H:H','I:I','J:J','K:K','L:L','M:M','N:N','O:O','P:P','Q:Q','R:R','S:S','T:T']
location_width = [80, 80, 10, 40, 28, 20, 20, 40, 60, 30, 60, 10, 15, 16, 26, 14, 14, 16, 32, 25]
rows = 1
opera_time = str(starttime.year) + 'y' + str(starttime.month) + 'm' + str(starttime.day) + 'd' + str(starttime.hour) + 'h' + str(starttime.minute) + 'm' + str(starttime.second) + 's'
filename = 'Springer_' + opera_time + '.xlsx'
workbook = xlsxwriter.Workbook(filename)  
worksheet = workbook.add_worksheet()
format_title=workbook.add_format()
format_title.set_border(1) 
format_title.set_bg_color('#cccccc') 
format_title.set_align('center') 
format_title.set_bold() 

worksheet.write_row('A1', parameter_name_list,format_title)
for i in xrange(20):
    worksheet.set_column(location_list[i], location_width[i])
for i in range(i_start,page_num):
    search_page_num = str(i + 1)
    search_url = 'http://link.springer.com/search/page/' + search_page_num + '?facet-discipline=%22'\
            + discipline + '%22&facet-language=%22' + Language_code + '%22&' + '&facet-content-type=%22'\
            + Content_type + '%22'
    search_page = get_data(search_url)
    search_page = "".join(search_page.split())
    search_info = re.findall('class="content-item-list">(.*?)</ol>',search_page)
    search_info = re.findall('no-accesshas-cover">(.*?)</li>',search_page)

    if i == i_start:
        j = j_start
    else:
        j = 0

    if i == page_num -1:
        j_end = last_item
    else:
        j_end = 20
    for j in range(j,j_end):
        last_item
        defstarttime = datetime.datetime.now()
        result_id = 'Page Num ' + str(i + 1) + ' & ' + str(j + 1) +'st Item' 
        print result_id + ' start!'
        book_doi_tmp = re.findall('ahref(.*?)tle=',search_info[j])
        book_doi = re.findall('="(.*?)"ti',book_doi_tmp[0])
        DOI = re.findall('/book/(.*?)"ti',book_doi_tmp[0])
        Content_Type = 'Book'
        if DOI == []:
            DOI = re.findall('/referencework/(.*?)"ti',book_doi_tmp[0])
            Content_Type = 'Reference Work'
        doi_urlcode = re.sub('/', '%2F',DOI[0])
        metrix_url = 'https://bookmetrix-proxy.live.cf.public.springer.com/books/' + doi_urlcode
        book_url = 'http://link.springer.com' + book_doi[0]
        book_page = get_data(book_url)
        book_page = re.sub('\n', '', book_page)
        about_this_book = re.findall('<h2>About this Book</h2>(.*?)<h3>Continue',book_page)
        Book_Title = re.findall('id="abstract-about-title">(.*?)</dd> ',about_this_book[0])
        Book_Subtitle = re.findall('id="abstract-about-book-subtitle">(.*?)</dd> ',about_this_book[0])
        if Book_Subtitle == []:
            Book_Subtitle = ['']
        Copyright = re.findall('id="abstract-about-book-chapter-copyright-year">(.*?)</dd> ',about_this_book[0])
        Print_ISBN = re.findall('id="abstract-about-book-print-isbn">(.*?)</dd> ',about_this_book[0])
        Online_ISBN = re.findall('id="abstract-about-book-online-isbn">(.*?)</dd> ',about_this_book[0])
        Publisher = re.findall('id="abstract-about-publisher">(.*?)</dd> ',about_this_book[0])
        Copyright_Holder = re.findall('id="abstract-about-book-copyright-holder">(.*?)</dd> ',about_this_book[0])
        Authors_info = re.findall('<li itemprop="editor"(.*?)</dd>',about_this_book[0])
        
        if Authors_info != []:
            Authors_info = re.findall('name">(.*?)</a>',Authors_info[0])
            Authors_info_num = len(Authors_info)
            for k in xrange(Authors_info_num):
                Authors_info[k] = Authors_info[k] + ' [editor]'
        else:
            Authors_info = re.findall('<li itemprop="author"(.*?)</dd>',about_this_book[0])
            Authors_info = re.findall('name">(.*?)</a>',Authors_info[0])
            
        Authors_info_num = len(Authors_info)
        Authors = ''
        if Authors_info_num > 1:
            for l in xrange(Authors_info_num):
                Authors = Authors + Authors_info[l] + '(' + str(l + 1) + ') '
        
        try:
            metrix_page = get_data(metrix_url)        

            print 'Metrix Data Exist & Saving The Items!'
            metrix_page = get_data(metrix_url)
            metrix_page = "".join(metrix_page.split())
            citation_count = re.findall('"citation_count":(.*?),',metrix_page)
            altmetric_mention_count = re.findall('"altmetric_mention_count":(.*?),',metrix_page)
            reader_count = re.findall('"reader_count":(.*?),',metrix_page)
            review_count = re.findall('"review_count":(.*?),',metrix_page)    
            download_count = re.findall('"download_count":(.*?)}',metrix_page)
            
            print Book_Title[0] + ' has saved successful !'

        except:
            f = open("error_log.txt",'a')  
            traceback.print_exc(file=f)
            f.flush()
            f = f.write('error: ' + '\n')
            
            print 'Metrix Data Do Not Exist & Will Try The Next One!'
            citation_count = ['No data']
            altmetric_mention_count = ['No data']
            reader_count = ['No data']
            review_count = ['No data']
            download_count = ['No data']

        timer = datetime.datetime.now()
        query_time = str(timer.year) + '(y)' + str(timer.month) + '(m)' + str(timer.day) + '(d)' + str(timer.hour) + '(h)' + str(timer.minute) + '(m)' + str(timer.second) + '(s)'

        parameter_list = [Book_Title[0], Book_Subtitle[0], Copyright[0], Authors, DOI[0], Print_ISBN[0], Online_ISBN[0], Publisher[0], Copyright_Holder[0], Discipline, book_url, Language, Content_Type, citation_count[0], altmetric_mention_count[0], reader_count[0], review_count[0], download_count[0], query_time, result_id]

        worksheet.set_row(rows,16) 
        for x in xrange(20):
            worksheet.write(rows, x, parameter_list[x])
        
        rows += 1
        defendtime = datetime.datetime.now()
        definterval = (defendtime - defstarttime).seconds
        usedinterval = (defendtime - starttime).seconds
        print str(Book_Title[0]) +'`s time is '+str(definterval/60)+' min ('+str(definterval)+' s )'
        print str(Book_Title[0]) +' Used time is '+str(int(usedinterval/3600))+' h '+str(int(usedinterval/60-int(usedinterval/3600)*60))+' m '+str(usedinterval-int(usedinterval/3600)*3600-(int(usedinterval/60-int(usedinterval/3600)*60))*60)+' s'+ '\n'
workbook.close() 
endtime = datetime.datetime.now()
interval = (endtime - starttime).seconds
print 'Total time is '+str(interval/60)+' min ('+str(interval)+' s )'



















#bold = workbook.add_format({'bold': True})    #定义一个加粗的格式对象
#worksheet.write('A'+str(rows), 'Hello')    #A1单元格写入'Hello'
#worksheet.write('A2', 'World', bold)    #A2单元格写入'World'并引用加粗格式对象bold
#worksheet.write('B2', u'中文测试', bold)    #B2单元格写入中文并引用加粗格式对象bold

        #保存excel           
#        with open('Springer.csv', 'wb') as csvfile:
#            spamwriter = csv.writer(csvfile,dialect='excel')
#            spamwriter.writerow(['Book_Title', 'Book_Subtitle', 'Copyright', 'Authors', 'DOI', 'Print_ISBN', 'Online_ISBN', 'Publisher', 'Copyright_Holder', 'Discipline', 'book_url', 'Language', 'Content_Type', 'citation_count', 'altmetric_mention_count', 'reader_count', 'review_count', 'download_count', 'query_time'])
#            spamwriter.writerow([Book_Title[0], Book_Subtitle[0], Copyright[0], Authors, DOI[0], Print_ISBN[0], Online_ISBN[0], Publisher[0], Copyright_Holder[0], Discipline, book_url, Language, Content_Type, citation_count[0], altmetric_mention_count[0], reader_count[0], review_count[0], download_count[0], query_time])
