import os
import json
import datetime
import win32com.client as win32

#ruta json data
path_json = r'C:\Users\Ronald\Desktop\client.json'

with open(path_json) as file_json:
    data = json.load(file_json)

client = data["client"]
dni = data["dni"]
type_credit = data["type credit"]

get_date_now = datetime.datetime.now()
get_hour_now = "{:02d}".format(get_date_now.hour)
get_minute_now = "{:02d}".format(get_date_now.minute)
get_day_now = get_date_now.day
get_month_now = "April"

if 0 <= get_date_now.hour <= 12:
    get_meridian_now = " AM"
else:
    get_meridian_now = " PM"

date = str(get_day_now)+ " de " + str(get_month_now)
hour = str(get_hour_now) + ":"+ str(get_minute_now) + get_meridian_now
print(date, hour)

word = win32.Dispatch('Word.Application')
word.Visible = True

#ruta modelo doc
doc = word.Documents.Open(r'C:\Users\Ronald\Desktop\_automate_doc_no_adeudo\formato\Constancia-no-adeudo.docx')

#reemplaza campos
range = doc.Content
range.Find.Execute('(titular)', False, False, False, False, False, True, 1, True, client, 2)
range.Find.Execute('(dni)', False, False, False, False, False, True, 1, True, dni, 2)
range.Find.Execute('(tipo_credito)', False, False, False, False, False, True, 1, True, type_credit, 2)
range.Find.Execute('(hora)', False, False, False, False, False, True, 1, True, hour, 2)
range.Find.Execute('(fecha)', False, False, False, False, False, True, 1, True, date, 2)

partial_path = r"C:\Users\Ronald\Desktop\_automate_doc_no_adeudo\generate_doc"
filename_doc = type_credit +"_"+client+".docx"
filename_pdf = type_credit +"_"+client+".pdf"
path_doc = os.path.join(partial_path, filename_doc)
path_pdf = os.path.join(partial_path, filename_pdf)

doc.SaveAs(path_doc)
doc.SaveAs(path_pdf, FileFormat=17)
doc.Close()

word.Quit()