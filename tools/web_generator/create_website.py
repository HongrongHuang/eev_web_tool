import os
import sys

class website:
    def __init__(self, current_html):
        self.current_html = current_html
        webfile = open(self.current_html, 'r')
        self.content = webfile.readlines()
        print "current html file:"
        print self.content
        for line in self.content:
            print line

    def add_new_line_with_text(self, text):
        self.content.append('<br>%s</br>\n'%(text))

    def get_all_content(self):
        return self.content

    def set_all_content(self, new_content):
        self.content = new_content

    def del_all_content(self):
        self.content = []

    def save_change(self):
        print "current html file:"
        for line in self.content:
            print line
        new_file = open(self.current_html, 'w')
        new_file.writelines(self.content)

def test():
    newweb = website(r'D:\workspace\evv_web_tools\xlrdT3a.py')
    file = newweb.get_all_content()
    new_content = []
    for line in file:
        new_content.append(line[4:])
	newweb.set_all_content(new_content)
    #newweb.save_change()

if __name__ == "__main__":

    print "hello world!"
    newweb = website(r'D:\workspace\web\index.html')
    
    import process_testresults
    new_xl = process_testresults.testresults(r'\\10.170.2.9\file_share_vol\D03_EE&Info&Con&Cloud\3.Electrical & Electronics\3.5 EEV\3.5.5 Labcar\04_TestCase\FRS5.0\FT\Test Case\M31T FRS5.0 System Test Report_20170508.xlsx')
    #new_xl = process_testresults.testresults(r'D:\report.xlsx')
    result_data = new_xl.get_all_importent_data()
    
    for item in result_data:
        newweb.add_new_line_with_text(item)
    newweb.save_change()
    #1/0
    #test()