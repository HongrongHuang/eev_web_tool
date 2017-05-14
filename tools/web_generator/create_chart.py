import os
import sys

def create_valid_link(path):
    # example: path = r'\\192.168.1.102\eclipse_workspace\'
    # output: path = "file://///192.168.1.102/eclipse_workspace/"
    result_path = path
    if path.startswith(r'\\'):
        result_path = result_path.replace('\\','/')
        result_path = "file:///" + result_path
    return result_path

class website:
    def __init__(self, current_html):
        self.current_html = current_html
        self.data_source_path = None
        webfile = open(self.current_html, 'r')
        self.content = webfile.readlines()
        print "current html file:"
        print self.content
        for line in self.content:
            print line
    def set_data_source_path(self, source_path):
        self.data_source_path = source_path
        basename = os.path.basename(self.data_source_path)
        dirname = os.path.dirname(self.data_source_path)
        for idx, line in enumerate(self.content):
            #print line
            if '{%sourcefilepathandname%}' in line:
                print "found line for {%sourcefilepathandname%}"
                new_content = create_valid_link(source_path)
                self.content[idx] = self.content[idx].replace('{%sourcefilepathandname%}', new_content)
                
            if '{%sourcefilename%}' in line:
                print "found line for {%sourcefilename%}"
                new_content = basename
                self.content[idx] = self.content[idx].replace('{%sourcefilename%}', new_content)
                
            if '{%sourcepathlink%}' in line:
                print "found line for {%sourcepathlink%}"
                new_content = create_valid_link(dirname)
                self.content[idx] = self.content[idx].replace('{%sourcepathlink%}', new_content)
                
            if '{%sourcepath%}' in line:
                print "found line for {%sourcepath%}"
                new_content = dirname
                self.content[idx] = self.content[idx].replace('{%sourcepath%}', new_content)
                
    
    def modify_table_data(self, chart_data):
        table_head = '<table border="1">\n'
        table_end = '</table>\n'
        
        row_head = '<tr>\n'
        row_head_red = '<tr bgcolor="#F02000">\n'
        row_head_gray = '<tr bgcolor="#CCCCCC">\n'
        row_end = '</tr>\n'
        
        cell_head = '<td>'
        cell_end = '</td>\n'
        
        
        for idx, line in enumerate(self.content):
            #print line
            if '{%htmltable%}' in line:

                htmltablestring = ''
                htmltablestring += table_head
                
                htmltablestring += row_head
                htmltablestring += cell_head + 'Function Group' + cell_end
                htmltablestring += cell_head + 'Pass' + cell_end
                htmltablestring += cell_head + 'Fail' + cell_end
                htmltablestring += cell_head + 'Not Test' + cell_end
                htmltablestring += cell_head + 'None' + cell_end
                htmltablestring += cell_head + 'Remark' + cell_end
                htmltablestring += row_end
                
                for item in chart_data:
                    print item
                    if item['Fail'] != 0:
                        htmltablestring += row_head_red
                    elif item['Not Test'] != 0:
                        htmltablestring += row_head_gray
                    else:
                        htmltablestring += row_head
                    
                    htmltablestring += cell_head + str(item['Function Group']) + cell_end
                    htmltablestring += cell_head + str(item['Pass']) + cell_end
                    htmltablestring += cell_head + str(item['Fail']) + cell_end
                    htmltablestring += cell_head + str(item['Not Test']) + cell_end
                    htmltablestring += cell_head + str(item['None']) + cell_end
                    htmltablestring += cell_head + str(item['Remark']) + cell_end
                    
                    htmltablestring += row_end
                htmltablestring += table_end
                
                new_content = str(htmltablestring)
                self.content[idx] = self.content[idx].replace('{%htmltable%}', new_content)
                
    def modify_chart_data(self, chart_data):
        for idx, line in enumerate(self.content):
            #print line
            if '{%height%}' in line:
                print "found line for height"
                item_number = len(chart_data)
                new_content = str(50*item_number)
                self.content[idx] = self.content[idx].replace('{%height%}', new_content)

            if '{%chart_data%}' in line:
                print "found line for data"
                
                temp_list_name = []
                temp_list_pass = []
                temp_list_fail = []
                temp_list_nottest = []
                for item in chart_data:
                    temp_list_name.append(item['Function Group'])
                    temp_list_pass.append(item['Pass'])
                    temp_list_fail.append(item['Fail'])
                    temp_list_nottest.append(item['Not Test'])
                new_data_in_list = [temp_list_name, temp_list_pass, temp_list_fail, temp_list_nottest]
                
                new_content = str(new_data_in_list)
                
                self.content[idx] = self.content[idx].replace('{%chart_data%}', new_content)
                
        
    def add_new_line_with_text(self, text):
        self.content.append('<br>%s</br>\n'%(text))

    def get_all_content(self):
        return self.content

    def set_all_content(self, new_content):
        self.content = new_content

    def del_all_content(self):
        self.content = []

    def save_change(self, new_file = None):
        print "current html file:"
        for line in self.content:
            pass
            #print line
        if new_file == None:
            new_file = open(self.current_html, 'w')
        else:
            new_file = open(new_file, 'w')
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
    
    chart_html_template = os.path.join(os.path.dirname(__file__), 'template', 'chart.html')
    newweb = website(chart_html_template)
    
    import process_testresults
    excelpath = r'\\CTCS-NOT-0118\huanghongrong\M31T FRS5.0 System Test Report_20170511.xlsx'
    #excelpath = r'\\10.170.2.9\file_share_vol\D03_EE&Info&Con&Cloud\3.Electrical & Electronics\3.5 EEV\3.5.5 Labcar\04_TestCase\FRS5.0\FT\Test Case\M31T FRS5.0 System Test Report_20170511.xlsx'
    excelpath = r'D:\report.xlsx'
    excelpath = r'\\192.168.1.102\eclipse_workspace\report.xlsx'
    excelpath = r'\\192.168.1.102\eclipse_workspace\M31T FRS5.0 System Test Report_20170511.xlsx'
    new_xl = process_testresults.testresults(excelpath)
    result_data, result_data_detail = new_xl.get_all_importent_data()
    
    newweb.modify_chart_data(result_data)
    newweb.modify_table_data(result_data_detail)
    
    newweb.set_data_source_path(excelpath)
    #newweb.save_change(r'D:\workspace\web\index.html')
    newweb.save_change(r'D:\workspace\web\2.html')
    #1/0
    #test()