import pandas as pd
import PySimpleGUI as sg


layout = [[sg.Text('Filename')],[sg.Input(), sg.FileBrowse()], [sg.OK(), sg.Cancel()]] 
window = sg.Window('File Splitter', layout)

event, file = window.Read()

def reader(file):
    dt =  pd.read_excel(file)
    columns_to_keep =['Full Name',	'PID','Mjr/Min/Cert', 'Mjr/Min/Cert Code','Dept/School','Mailing Address 1','Mailing Address City',
 'Mailing Address State', 'Mailing Address Zip','Advisor(s)']
    
    dt=dt[columns_to_keep]
    dt.columns = ['Name','PID', 'Major','Major_Code','Dept', 'Address','City','State','Zip','Advisor(s)']

    phd_codes =  dt.Major_Code.str.contains('^PH|^ED')
    sports_codes = dt.Dept.str.contains('Recreation and Sport Pedagogy')

    phd = dt[phd_codes]
    sports =  dt[sports_codes]
    rest =  dt[~phd_codes & ~sports_codes]

    with pd.ExcelWriter('Output.xlsx', engine='xlsxwriter') as writer:
        dt.to_excel(writer, sheet_name='All', index=False)
        phd.to_excel(writer,sheet_name= 'PHD', index=False)
        rest.to_excel(writer,sheet_name='Masters', index=False)
        sports.to_excel(writer, sheet_name ='Recreation', index=False)

fname = file[0]

if fname:
    try:
        reader(fname)
        sg.Popup("Done!\nThe file Output.xlsx is saved in the same location.",title ="Completed")
    except KeyError:
        sg.Popup("Error!\nColumn(s) may need to be renamed or they do not exist.\nPlease check the file structure!", title='Error')
    except Exception:
        sg.Popup('Unsupported file format. Only Excel files!',title="Error")
