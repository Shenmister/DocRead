import os
import xlsxwriter
import extract_msg


#format columns
def filewrite(filename):
    f = r'%s'%filename
    msg = extract_msg.Message(f)
    msg_sender = msg.sender
    msg_date = msg.date
    msg_subj = msg.subject
    msg_CC = msg.cc
    msg_To = msg.to
    msg_message = msg.body
    if len(msg_message) > 400:
        msg_message = msg_message[:400]
    
    index = msg_date.index(":")
    msg_Time = msg_date[index-2:index+6]
    msg_date = msg_date[5:index-3]

    msg_hyper = os.path.abspath(filename)
    msg_Rename = 'X1331'
    file_prop = [msg_To, msg_sender, msg_CC,msg_Time,msg_date,msg_subj, msg_message,msg_hyper,msg_Rename]
    return file_prop




def main():
    workbook = xlsxwriter.Workbook('sorted.xlsx')
    worksheet = workbook.add_worksheet()

    data_cols = ['To', 'From', 'CC', 'Time', 'Date', 'Title of Email','Description (Content of Email)','Hyperlink to file','Renamed As']
    header_format = workbook.add_format({
        'bold': False,
        'font_name': 'Arial',
        'font_size': 10,
        'text_wrap': True,
        'center_across': True,
        'valign': 'bottom',
        'fg_color': '#cdffff',
        'border': 1})
    #setting up columns
    col_num1 = 0
    for i in range(len(data_cols)):
        worksheet.write(0,col_num1, data_cols[i],header_format )
        col_num1+=1

    #temp storage 
    files = []
    
    #reading files
    directory_str = str(os.path.abspath(os.getcwd()))
    directory = os.fsencode(directory_str)    
    for file in os.listdir(directory):
        filenm = os.fsdecode(file)
        if filenm.endswith(".msg"):
            files.append(filewrite(filenm))
        else: continue

    #filling spreadsheett
    row2 = 1
    for i in files:
        row2 +=1 
        col2 = 0
        for j in i:
            if col2 == 7: worksheet.write_url(row2,col2,j)
            else: worksheet.write(row2,col2,j)
            col2+=1


    

    workbook.close()

main()