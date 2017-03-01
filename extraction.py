from xlrd import *
class Extraction:
    def alphanumeric_list(self):
        list_alphnum=[]
        for i in xrange(26):
            list_alphnum.append(chr(i+65))
        for i in xrange(10):
            list_alphnum.append(str(i))
        return list_alphnum

    def capitalised(self,test):
        alphn_list=self.alphanumeric_list()
        capital=''
        try:
            for letter in test.upper():
                if letter in alphn_list:
                    capital+=letter
        except Exception as e:
            print e

           
        return capital
    def extract_data(self,excel_wb,sheet_name='',sheet_index=0):
         standard_l={}
         if sheet_name!='':
             sheet=excel_wb.sheet_by_name(sheet_name)
         else:
             sheet=excel_wb.sheet_by_index(sheet_index)
         row,col=sheet.nrows,sheet.ncols
         for col_in_sheet in range(col):
            temp=[]
            for row_in_sheet in range(row):
                if sheet.cell(row_in_sheet,col_in_sheet).value!='' and row_in_sheet!=0:
                    temp.append(sheet.cell(row_in_sheet,col_in_sheet).value)
            standard_l[sheet.cell(0,col_in_sheet).value]=temp
         return standard_l
    def search_titles(self,excel_sample,standard):
        cap=''
        sheetwise_cordinates={}
        for sheet in excel_sample.sheets():
            if sheet.visibility==0:
                cordinates=()
                print sheet.name
                for row_in_sheet in range(sheet.nrows):
                    for col_in_sheet in range(sheet.ncols):
                        if sheet.cell(row_in_sheet,col_in_sheet).ctype==XL_CELL_TEXT:
                            cap=self.capitalised(sheet.cell(row_in_sheet,col_in_sheet).value)
                            if  cap in str(standard.keys()) and len(cap)>1:
                                cordinates=cordinates+((row_in_sheet,col_in_sheet),)
                sheetwise_cordinates[sheet.name]=cordinates
        return sheetwise_cordinates
    def extract_rows(self,cordinates_data):
        row_data={}
        l=[]
        for key in cordinates_data:
            for x,y in cordinates_data[key]:
                if x not in l:
                    l.append(x)
            row_data[key]=l
        return row_data
    def get_cordinates(self,cord,workbook):
        temp=()
        cords_d={}
        cords_dict={}
        for val in cord:
            l=[]
            for data in cord[val]:
                temp=()
                sheet=workbook.sheet_by_name(val)
                if data<sheet.nrows:
                    for col in range(sheet.ncols):
                        if col<sheet.ncols-1 and data<sheet.nrows-1 and  sheet.cell(data,col).value!='' and sheet.cell(data,col+1).value=='' and sheet.cell(data+1,col).value!='' :
                            temp+=((sheet.cell(data+1,col).value),col),
                        elif col<sheet.ncols-1 and data<sheet.nrows-1 and sheet.cell(data,col).value==''  :
                            if sheet.cell(data,col+1).value!='' and sheet.cell(data+1,col).value!='':
                                temp+=((sheet.cell(data+1,col).value),col),
                            elif sheet.cell(data,col+1).value=='' and sheet.cell(data+1,col).value!='':
                                temp+=((sheet.cell(data+1,col).value),col),
                            elif sheet.cell(data,col+1).value=='' and sheet.cell(data-1,col).value!='':
                                temp+=((sheet.cell(data-1,col).value),col),
                            elif sheet.cell(data,col+1).value!='' and sheet.cell(data-1,col).value!='':
                                temp+=((sheet.cell(data-1,col).value),col),
                        else:
                            if col<sheet.ncols-1 and data<sheet.nrows-1 and sheet.cell(data,col).value!='':
                                temp+=((sheet.cell(data,col).value),col),
                    if temp!='' and len(temp)>1:
                        cords_dict[data]=temp
            cords_d[val]=cords_dict
        return cords_d
