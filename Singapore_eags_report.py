import os
import re
import sys
import time
import glob
import logging
from mail_alert import send_mail
import pandas as pd
import xlwings as xw
import xlwings.constants as win32c
from datetime import date, datetime
from PIL import ImageGrab
import bu_alerts


drive = r"K:\_Credit Calc\Hamilton Metals Credit Report\AR Credit Report Automation"

def remove_existing_files(files_location):
    """_summary_

    Args:
        files_location (_type_): _description_

    Raises:
        e: _description_
    """           
    logging.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logging.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")
        print("Pause")
    except Exception as e:
        logging.exception("Exception in: remove_existing_files()")
        logging.exception(e)
        raise e


def freezepanes_for_tab(cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Rows(cellrange).Select()
        working_workbook.app.api.ActiveWindow.FreezePanes = True
    except Exception as e:
        raise e      
    

def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e
    

def xlOpner(inputFile):
    try:
        retry = 0
        while retry<10:
            try:
                input_wb = xw.Book(inputFile, update_links=False)
                return input_wb
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
    except Exception as e:
        print(f"Exception caught in xlOpner :{e}")
        logging.info(f"Exception caught in xlOpner :{e}")
        raise e
    

def remove_special_characters(my_pdf,column_list):
    try:
        # column_list = list(my_pdf.columns[[-5,-4,-3,-2]])
        logging.info("inside remove special characters")
        for values in column_list:
            my_pdf[values] = my_pdf[values].astype(str)
            my_pdf[values]  = [x[values].replace('$', '') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace('(', '-') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace(')', '') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace(',', '') for i, x in my_pdf.iterrows()]
            # my_pdf[values]  = [x[values].replace('0.0', '0.00') for i, x in my_pdf.iterrows()]
            my_pdf[values] = my_pdf[values].astype(float)
            # my_pdf[values]  = [x[values].replace('0.00', '0') for i, x in my_pdf.iterrows()]
            
        return  my_pdf   
    except Exception as e:
        raise e  


def interior_coloring_temp(colour_value,cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Range(cellrange).Select()
        a = working_workbook.app.selection.api.Interior
        a.Pattern = win32c.Constants.xlSolid
        a.PatternColorIndex = win32c.Constants.xlAutomatic
        a.Color = colour_value
        a.TintAndShade = 0
        a.PatternTintAndShade = 0        
    except Exception as e:
        raise e  


def insert_top1_btm2_borders(cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Range(cellrange).Select()
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeLeft).LineStyle = win32c.Constants.xlNone
        # linestylevalues=[win32c.BordersIndex.xlEdgeLeft,win32c.BordersIndex.xlEdgeTop,win32c.BordersIndex.xlEdgeBottom,win32c.BordersIndex.xlEdgeRight,win32c.BordersIndex.xlInsideVertical,win32c.BordersIndex.xlInsideHorizontal]
        # for values in linestylevalues:
        a=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeTop)
        a.LineStyle = win32c.LineStyle.xlContinuous
        a.ColorIndex = 0
        a.TintAndShade = 0
        a.Weight = win32c.BorderWeight.xlThin
        b=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeBottom)
        b.LineStyle = win32c.LineStyle.xlDouble
        b.ColorIndex = 0
        b.TintAndShade = 0
        b.Weight = win32c.BorderWeight.xlThick
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeRight).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideVertical).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideHorizontal).LineStyle = win32c.Constants.xlNone
    except Exception as e:
        raise e
    

def row_range_calc(filter_col:str, input_sht,wb):
    sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row

    sp_address= input_sht.api.Range(f"{filter_col}2:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address

    sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

    row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

    while row_range[-1]!=sp_lst_row:

        sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row

        sp_address = sp_address+','+(input_sht.api.Range(f"{filter_col}{row_range[-1]+1}:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)

        # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

        row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
        
    
    sp_address = sp_address.replace("$","").split(",")
    init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
    sublist = []
    flat_list = [item for sublist in init_list for item in sublist]
    return flat_list, sp_lst_row,sp_address


    
def proceesing_report(template_wb,raw_wb,compare_wb):
    try:  
        initial_sheet =  raw_wb.sheets[0]  
        try:
            template_wb.sheets['Raw Data'].delete()
            initial_sheet.api.Move(Before=template_wb.api.Sheets(3))
            raw_data_sheet = template_wb.sheets[2] 
            raw_data_sheet.name= 'Raw Data'

            ######### formatting sheet(deleting columns) #####################

            lst_rw_raw = raw_data_sheet.api.UsedRange.Rows.Count
            raw_data_sheet.range(f"H:J").api.Delete()
            raw_data_sheet.autofit()

            ######### Clearing Sheet #####################

            data_sheet = template_wb.sheets['Data']
            clr_rw_rd = data_sheet.api.UsedRange.Rows.Count
            data_sheet.activate()
            data_sheet.range(f"3:{clr_rw_rd}").api.EntireRow.Delete()

            ############# Updating Raw Data Sheet #############

            raw_data_sheet.api.Range(f"B2:H{lst_rw_raw}").Copy()
            data_sheet.api.Range(f"B2")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)            
            template_wb.app.api.CutCopyMode=False

            ########## converting date and dragging formulas ###########

            data_sheet.range(f"A3").value = "2"
            data_sheet.api.Range(f"A2:A3").Select()
            # data_sheet.api.Range(f"A2:A{lst_rw_raw}").FillDown()
            template_wb.app.api.Selection.AutoFill(Destination=data_sheet.api.Range(f"A2:A{lst_rw_raw}"),Type=win32c.AutoFillType.xlFillSeries)
            data_sheet.autofit()

            ############### applying sums ##################

            data_sheet.range(f"H{lst_rw_raw+1}").value = f"=SUM(H2:H{lst_rw_raw})"
            insert_top1_btm2_borders(cellrange=f"H{lst_rw_raw+1}",working_sheet=data_sheet,working_workbook=template_wb)
            interior_coloring_temp(16776960,cellrange=f"H{lst_rw_raw+1}",working_sheet=data_sheet,working_workbook=template_wb)
            data_sheet.api.Range(f"H{lst_rw_raw+1}").Select()
            template_wb.app.api.Selection.AutoFill(Destination=data_sheet.api.Range(f"H{lst_rw_raw+1}:R{lst_rw_raw+1}"),Type=win32c.AutoFillType.xlFillDefault)
            data_sheet.autofit()
            interior_coloring_temp(10092288,cellrange=f"H1:H{lst_rw_raw}",working_sheet=data_sheet,working_workbook=template_wb)
            data_sheet.api.Range(f"H{lst_rw_raw+1}:R{lst_rw_raw+1}").Font.Bold = True

            ############# Pivot Refresh and data count ########################

            pivot_sheet = template_wb.sheets['Pivot']
            pivot_sheet.activate()
            num_col = data_sheet.range('A1').end('right').column        
            pivotCount = template_wb.api.ActiveSheet.PivotTables().Count
            for j in range(1, pivotCount+1):
                if template_wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData != f"'Data'!R3C1:R{lst_rw_raw}C{num_col}": #Updateing data source
                    template_wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'Data'!R1C1:R{lst_rw_raw}C{num_col}" #Updateing data source
                template_wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()
            count = pivot_sheet.range(f"A4").expand('down').count - 1
            
            ################ refreshing data on eags sgp sheet ############

            Eags_sgp_sheet = template_wb.sheets['EAGS_SGP']
            Eags_sgp_sheet.activate()
            Eags_sgp_sheet.api.Range(f"A4:L4").Select()
            template_wb.app.api.Selection.AutoFill(Destination=Eags_sgp_sheet.api.Range(f"A4:L{count+3}"),Type=win32c.AutoFillType.xlFillDefault)

            lst_row = Eags_sgp_sheet.range(f'A'+ str(Eags_sgp_sheet.cells.last_cell.row)).end('up').row
            pivot_sheet.api.Range(f"A4:A{lst_row}").Copy()
            Eags_sgp_sheet.api.Range(f"A4")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)  
            template_wb.app.api.CutCopyMode=False 
            pivot_sheet.api.Range(f"B4:I{lst_row}").Copy()
            Eags_sgp_sheet.api.Range(f"E4")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues) 
            template_wb.app.api.CutCopyMode=False 
            Eags_sgp_sheet.autofit()

            print("proceess completed for xcel")

            lst_row2 = Eags_sgp_sheet.range(f'A'+ str(Eags_sgp_sheet.cells.last_cell.row)).end('up').row    
            Eags_sgp_sheet.range(f"A4:L{lst_row2}").api.Sort(Key1=Eags_sgp_sheet.range(f"E4:E{lst_row2}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)    

            ############### New customer Logic #############

            customer_list = []
            value_list = Eags_sgp_sheet.range(f"D4:D{lst_row2}").value
            for index,payemnt_terms in enumerate(value_list):
                if payemnt_terms==None:
                    print(f'New customer Found::{Eags_sgp_sheet.range(f"A{index+4}").value}')
                    customer_list.append(Eags_sgp_sheet.range(f"A{index+4}").value)
                    time.sleep(2)
                    Eags_sgp_sheet.range(f"A{index+4}:L{index+4}").api.CopyPicture(Appearance=1, Format=2)
                    time.sleep(2)
                    image_name = f"new_customer_{index+4}"
                    # grab the saved image from the clipboard and save to working directory
                    failure_image_path = drive + f'\\EAGS SINGAPORE REPORT\\Failure_Uploads\\{image_name}.png'
                    time.sleep(1)
                    ImageGrab.grabclipboard().save(failure_image_path)
                    time.sleep(1)
                    locations_list.append(failure_image_path)
                    continue
                else:
                    print("No new customers found today")
    
            if len(glob.glob(raw__path__+"\\*.png"))>0:
                print("ending the process")
                try:
                    template_wb.app.quit()
                except:
                    pass
                receiver_email = "yashn.jain@biourja.com,arun.kaul@biourja.com,pravin.anthon@biourja.com,neeraj.gupta@biourja.com,bharat.pathak@biourja.com"
                # receiver_email = "yashn.jain@biourja.com"
                nl = '<br>'

                
                customers_html = "<ul>" + "".join([f"<li>{customer}</li>" for customer in customer_list]) + "</ul>"
                body = (f'{nl}<strong>New Customer Notification</strong>,{nl}New Customers:{customers_html} {nl}{nl}Action:Please add these new customers to the <strong>STX_SGP Sheet.</strong>{nl}Attached path for the excel=<u>{template_workbook}</u>{nl}')
                send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully,{body}',multiple_attachment_list = locations_list)
                sys.exit(0)     

            ########## moving on to comparision sheet #####################

            compare_wb.sheets.add(f"IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}",after=compare_wb.sheets['Sheet2'])
            it_eags_sgp_sheet = compare_wb.sheets[f"IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}"]
            Eags_sgp_sheet.api.UsedRange.Copy()
            it_eags_sgp_sheet.range('A1').paste()
            template_wb.app.api.CutCopyMode=False
            it_eags_sgp_sheet.autofit()

            ############# moving onto sheet2 ########

            previous_day_sheet = compare_wb.sheets[3].name
            sheet2 = compare_wb.sheets[f"Sheet2"]
            sheet2.activate()
            sheet2.range(f"C4").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!E3"
            sheet2.range(f"C5").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!F3"
            sheet2.range(f"C6").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!G3"
            sheet2.range(f"C7").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!H3"
            sheet2.range(f"C8").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!I3"
            sheet2.range(f"C9").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!J3"
            sheet2.range(f"C10").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!K3"
            sheet2.range(f"C11").value = f"='IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!L3"


            ############ updating yesterdays values ###########

            sheet2.range(f"D4").value = f"='{previous_day_sheet}'!E3"
            sheet2.range(f"D5").value = f"='{previous_day_sheet}'!F3"
            sheet2.range(f"D6").value = f"='{previous_day_sheet}'!G3"
            sheet2.range(f"D7").value = f"='{previous_day_sheet}'!H3"
            sheet2.range(f"D8").value = f"='{previous_day_sheet}'!I3"
            sheet2.range(f"D9").value = f"='{previous_day_sheet}'!J3"
            sheet2.range(f"D10").value = f"='{previous_day_sheet}'!K3"
            sheet2.range(f"D11").value = f"='{previous_day_sheet}'!L3"            

            ########## moving on to compare sheet ##################
            
            compare_sheet = compare_wb.sheets[f"Compare"]
            compare_sheet.activate() 
            # column_list = compare_sheet.range("A1").expand('right').value
            total_ar_present = 3
            list2=[f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,5,0),0)",f"=IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,5,0),0)",f"=C6-D6",\
                   f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,6,0),0)-IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,6,0),0)",\
                    f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,7,0),0)-IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,7,0),0)",\
                    f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,8,0),0)-IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,8,0),0)",\
                    f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,9,0),0)-IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,9,0),0)",\
                    f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,10,0),0)-IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,10,0),0)",\
                    f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,11,0),0)-IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,11,0),0)",\
                    f"=IFERROR(VLOOKUP(@$A:$A,'IT_EAGS_SGP {datetime.strftime(today_date,'%m.%d.%y')}'!$A:$M,12,0),0)-IFERROR(VLOOKUP(@$A:$A,'{previous_day_sheet}'!$A:$M,12,0),0)"]
            
            last_row = compare_sheet.range(f'A'+ str(compare_sheet.cells.last_cell.row)).end('up').row
            for values in list2:
                last_column_letter=num_to_col_letters(total_ar_present)
                compare_sheet.range(f"{last_column_letter}6").value = values
                time.sleep(1)
                compare_sheet.range(f"{last_column_letter}6").copy(compare_sheet.range(f"{last_column_letter}6:{last_column_letter}{last_row}"))
                total_ar_present+=1

            compare_sheet.range(f"A6:L{last_row}").api.Sort(Key1=compare_sheet.range(f"E6:E{last_row}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)    

            df = compare_sheet.range(f"A6:E{last_row}").options(pd.DataFrame,header=False,index=False).value

            increase_check = None
            top_increasing_customers = df[df[4]>0]
            if len(top_increasing_customers)>0:
                top_increasing_customers = top_increasing_customers[[0,4]]
                compare_sheet.range("O32").options(header=False,index=False).value = top_increasing_customers
                compare_sheet.range("P31").value = f"=SUM(P32:P{31+len(top_increasing_customers)})"
                increase_check= True
            else:
                print("no customers found under change>0")

            decrease_check = None
            top_decreasing_customers = df[df[4]<0]    
            if len(top_decreasing_customers)>0:
                top_decreasing_customers = top_decreasing_customers[[0,4]]
                compare_sheet.range("O5").options(header=False,index=False).value = top_decreasing_customers
                compare_sheet.range("P4").value = f"=SUM(P5:P{4+len(top_decreasing_customers)})"
                decrease_check= True
            else:
                print("no customers found under change<0")

            if increase_check:
                time.sleep(2)
                compare_sheet.range(f"O31:P{31+len(top_increasing_customers)}").api.CopyPicture(Appearance=1)
                time.sleep(2)
                # grab the saved image from the clipboard and save to working directory
                top_increasing_image_path = drive + "\\EAGS SINGAPORE REPORT" +"\\PNG Uploads" +"\\top_increasing.png"
                ImageGrab.grabclipboard().save(top_increasing_image_path)
                locations_list.append(top_increasing_image_path)
            else:
                time.sleep(2)
                compare_sheet.range(f"O31:P31").api.CopyPicture(Appearance=1)
                time.sleep(2)
                # grab the saved image from the clipboard and save to working directory
                top_increasing_image_path = drive + "\\EAGS SINGAPORE REPORT" +"\\PNG Uploads" +"\\top_increasing.png"
                ImageGrab.grabclipboard().save(top_increasing_image_path)
                locations_list.append(top_increasing_image_path)   

            if decrease_check:
                time.sleep(2)
                compare_sheet.range(f"O4:P{4+len(top_decreasing_customers)}").api.CopyPicture(Appearance=1)
                time.sleep(2)
                # grab the saved image from the clipboard and save to working directory
                top_decreasing_image_path = drive + "\\EAGS SINGAPORE REPORT" +"\\PNG Uploads" +"\\top_decreasing.png"
                ImageGrab.grabclipboard().save(top_decreasing_image_path)
                locations_list.append(top_decreasing_image_path)

            else:
                time.sleep(2)
                compare_sheet.range(f"O4:P4").api.CopyPicture(Appearance=1)
                time.sleep(2)
                # grab the saved image from the clipboard and save to working directory
                top_decreasing_image_path = drive + "\\EAGS SINGAPORE REPORT" +"\\PNG Uploads" +"\\top_decreasing.png"
                ImageGrab.grabclipboard().save(top_decreasing_image_path)
                locations_list.append(top_decreasing_image_path)                    

            sheet2.activate()
            time.sleep(2)
            sheet2.range(f"B2:E11").api.CopyPicture(Appearance=1,)
            time.sleep(2)
            # grab the saved image from the clipboard and save to working directory
            credit_report_image_path = drive + "\\EAGS SINGAPORE REPORT" +"\\PNG Uploads" +"\\credit_report.png"
            ImageGrab.grabclipboard().save(credit_report_image_path)

            time.sleep(2)
            sheet2.shapes[0].api.CopyPicture(Appearance=1,)
            time.sleep(2)
            # grab the saved image from the clipboard and save to working directory
            total_ar_outstanding_image_path = drive + "\\EAGS SINGAPORE REPORT" +"\\PNG Uploads" +"\\total_ar_outstanding.png"
            ImageGrab.grabclipboard().save(total_ar_outstanding_image_path)

            time.sleep(2)
            sheet2.shapes[1].api.CopyPicture(Appearance=1)
            time.sleep(2)
            # grab the saved image from the clipboard and save to working directory
            total_cr_past_image_path = drive + "\\EAGS SINGAPORE REPORT" +"\\PNG Uploads" +"\\total_cr_past.png"
            ImageGrab.grabclipboard().save(total_cr_past_image_path)


            html_body= """
            <style>
                img {
                    background-size: cover;
                    max-height: 50vh;
                    max-width: 70vh;
                    margin: 5px;
                }

                .class1 {
                    font-family: 'Lucida Sans', 'Lucida Sans Regular', 'Lucida Grande', 'Lucida Sans Unicode', Geneva, Verdana, sans-serif;
                    color: rgba(0, 0, 255, 0.062)44, 44, 110;
                }

                .class2 {
                    font-family: 'Courier New', Courier, monospace;
                    color: rgba(0, 0, 255, 0.062)44, 44, 110;
                }

                .img1 {
                    /* border-width: 10px solid black; */
                    border: 2px solid black;
                    /* box-shadow: -10px 0 10px rgba(128, 128, 128, 0.5); */
                }

                .top_right_bottom {
                    border-top: 2.5px solid black;
                    border-bottom: 2.5px solid black;
                    border-right: 2.5px solid black;
                    box-shadow: -14px 10px 12px 0px rgba(128, 128, 128, 0.5);
                }

                .top_right {
                    border-top: 2.5px solid black;
                    border-right: 2.5px solid black;
                    box-shadow: -14px 10px 12px 0px rgba(128, 128, 128, 0.5);

                }

                .box-shadow {
                    position: absolute;
                    top: 0;
                    left: -10px;
                    /* Adjust the offset to control the shadow position */
                    width: 10px;
                    /* Adjust the width of the shadow as needed */
                    height: 100%;
                    background: grey;
                    /* Color of the shadow */
                    z-index: -1;
                }

            h3{
            width: 100%;
            margin: 0;
            padding: 0;
            text-align: center;
            font-style: underline;
            }
            
            h3:after {
            display: inline-block;
            margin: 0 0 3px 20px;
            height: 3px;
            content: " ";
            text-shadow: none;
            background-color: #999;
            width: 50vh;
            }
            
            h3:before {
            display: inline-block;
            margin: 0 20px 3px 0;
            height: 3px;
            content: " ";
            text-shadow: none;
            background-color: #999;
            width: 50vh;
            }
            .class1:hover
            {
                opacity: 0.5;
            }
            u{
                margin-left: 8px;
            }
            </style>

            <body>


            <br>
            <c class="class1">

                Morning to All,
            </c>
            <br>
            <br>
            <c class="class1">
                Please find attached Energy Alloys Singapore AR Credit report.
            </c>
            <br>
            <br>
            <br>
            <!-- <h3 style="text-align: center;text-transform: capitalize; padding: 0;margin: 0;"> -->
            <h3 class="class2">
            <u>SINGAPORE</u>
            </h3>
            
            <!-- ====================================================================== -->
            <!-- <br> -->
            <br>
            <div style="display:flex; justify-content:center;flex-direction:row">
                <img alt="Embedded Image" class="top_right" src="credit_report.png" />
                <img alt="Embedded Image" class="img1" src="total_ar_outstanding.png" />
            </div>

            <br>

            <div style="display:flex; justify-content: center;flex-direction:row">

                <img alt="Embedded Image" class="img1" src="total_cr_past.png" />
                <!-- <img alt="Embedded Image" class="img1" src="total_ar_outstanding.png" /> -->
                <div style="display: flex; justify-content: space-evenly; flex-direction: column;margin-top: 7px;">

                    <b><u> Payments Received today:</u></b>
            

                    <img alt="Embedded Image" class="top_right_bottom" src="top_decreasing.png" />
        
                    <br>
                    <b><u> New Invoicing done today:</u></b>
            

                    <!-- <img alt="Embedded Image" src="top_increasing.png"/> -->
                    <img alt="Embedded Image" class="top_right_bottom" src="top_increasing.png" />
                    <br>

                </div>
            </div>
    
            <br>
            Thanks, and Regards.
            <br>
            <br>
            <p style="text-align:center">Copyright © 2023 IT India. All rights reserved. For additional support or queries, please email us at <strong><u>devsupport@biourja.com.</strong></u></p>              
            </body>"""
            stx_sgp_sheet = template_wb.sheets['STX_SGP'] 
            tablist={pivot_sheet:255,Eags_sgp_sheet:49407,data_sheet:15773696,raw_data_sheet:5287936,stx_sgp_sheet:65535}
            for tab,color in tablist.items():
                freezepanes_for_tab(cellrange="2:2",working_sheet=tab,working_workbook=template_wb) 
                try:
                    tab.api.Tab.Color = color
                except:
                    tab.api.Tab.ThemeColor =color
                tab.api.AutoFilterMode=False 

        except Exception as e:
            logging.exception(str(e))
            print("Error while generating Eags singapore sheet")
            raise e

        return html_body
    except Exception as e:
        raise e


if __name__ == "__main__":
    try:

        job_name="BIO_PAD01_TEST_EAGS_SINGAPORE_CREDIT_REPORT"
        # log progress --
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        logfile = os.getcwd() + '\\' + 'logs' + '\\' + f'{job_name}.txt'

        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logging.info("Execution Started")

        locations_list = []
        # logging.info('setting receiver_email')
        receiver_email = "yashn.jain@biourja.com,arun.kaul@biourja.com,pravin.anthon@biourja.com,neeraj.gupta@biourja.com,bharat.pathak@biourja.com"
        # receiver_email = "yashn.jain@biourja.com,imam.khan@biourja.com,apoorva.kansara@biourja.com, accounts@biourja.com, rini.gohil@biourja.com,itdevsupport@biourja.com"
        raw__path__ = drive + f'\\EAGS SINGAPORE REPORT\\Failure_Uploads'
        remove_existing_files(raw__path__)
        time_start=time.time()
        today_date=date.today()
        raw_file_path = drive + "\\EAGS SINGAPORE REPORT"+"\\Input"
        if len(glob.glob(raw_file_path+"\\*.xls"))>0:
            raw_file = glob.glob(raw_file_path+"\\*.xls")[0]    
            pathraw, file_name_inv = os.path.split(raw_file)
            # year = today_date.year
            # pre_month = int(re.findall("\d+",file_name_inv)[0]) - 1
            # pre_date = today_date.replace(month=pre_month)
            # today_date = today_date.replace(month=int(re.findall("\d+",file_name_inv)[0]))
            # pre_date_fldr = pre_date.strftime("%m-%y")
            # date_fldr = today_date.strftime("%m-%y")
            small_yr = today_date.strftime("%y")
        else:
            logging.info(f"Report not found ::: {raw_file_path}")   
            locations_list.append(logfile)
            send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully,Raw file not found here ::: {raw_file_path}',multiple_attachment_list = locations_list)
                 

        comaparision_workbook = drive + "\\EAGS SINGAPORE REPORT" + "\\Comparision Report" + f'\\Top 10 Increase & Decrease - Change 08.01.2023.xlsx'
        if not os.path.exists(comaparision_workbook):
            logging.info(f"{comaparision_workbook} Excel file not present")           

        template_workbook = drive + "\\EAGS SINGAPORE REPORT" + "\\Template File"+f'\\EAGS_SGP_ Credit Report_template.xlsx'
        if not os.path.exists(template_workbook):
            logging.info(f"{template_workbook} Excel file not present")


        try:
            raw_wb = xlOpner(raw_file)
        except Exception as e:
            logging.info(f"could not open workbook: {raw_file}")
            raise e
      
        raw_wb.api.AutoFilterMode=False
        raw_wb.app.api.CutCopyMode=False

        try:
            template_wb = xlOpner(template_workbook)
        except Exception as e:
            logging.info(f"could not open workbook: {template_workbook}")
            raise e
      
        template_wb.api.AutoFilterMode=False
        template_wb.app.api.CutCopyMode=False        

        try:
            compare_wb = xlOpner(comaparision_workbook)
        except Exception as e:
            logging.info(f"could not open workbook: {comaparision_workbook}")
            raise e
      
        compare_wb.api.AutoFilterMode=False
        compare_wb.app.api.CutCopyMode=False 

        try:
            html_body = proceesing_report(template_wb,raw_wb,compare_wb)
        except Exception as e:
            logging.info(f"Inbound/Outbound Tab Failure : {e}")
            raise e        
        print("Done")
        
        output_location = rf'{drive}\EAGS SINGAPORE REPORT\Output'
        if not os.path.exists(output_location):
            os.makedirs(output_location)

        try:
            template_wb.save(f"{output_location}\\EAGS_SGP_ Credit Report_{today_date}.xlsx")
            print(f"inventory done and saved in {output_location}")
            compare_wb.save(f"{output_location}\\EAGS_SGP_ Credit Report_{today_date}.xlsx")
            print(f"inventory done and saved in {output_location}")
            wb_name = template_wb.name
            template_wb.app.quit()
        except Exception as e:
            logging.info(f"could not save or kill ::: {wb_name}")
            raise e 
        
        time.sleep(2)
        # remove_existing_files(raw_file_path) #####please uncomment on prod ########
        logging.info(f"files succesfully removed from folder :::: {raw_file_path}")
        locations_list.append(logfile)
        locations_list.append(f"{output_location}\\EAGS_SGP_ Credit Report_{today_date}.xlsx")
        send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = html_body,multiple_attachment_list = locations_list)     


    except Exception as e:
        logging.exception(str(e))
        try:
            template_wb.app.quit()
        except:
            pass    
        send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)

    finally:
        logging.info("Process completed")
        print("process completed")







