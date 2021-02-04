import os
import sys
import pandas as pd
import pyexcel
import requests
import shutil
import datetime
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

EXCEL_FILE_PATH = ""
DIR_PATH = 'workdir/pdf'
columns = ['Дата сертификата', '№ сертификата', 'Номер т/с', 'Ссылка на загрузку сертификата в PDF']
HEADERS={
   "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 YaBrowser/20.12.2.108 Yowser/2.5 Safari/537.36",
   "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
   "Accept-Encoding" : "gzip, deflate, br",
   "Accept-Language" : "ru,en;q=0.9"
}

def pdf_dowload(url, name):

   filereq = requests.get(url, verify=False, stream=True, headers=HEADERS)
   with open(name,"wb") as receive:
         shutil.copyfileobj(filereq.raw,receive)
   del filereq


if __name__ == '__main__':
   pdfdirlist=os.listdir(os.path.dirname(__file__)+'\\pdf\\')
   for fpdf in pdfdirlist:
      os.remove(os.path.join(os.path.dirname(__file__)+'\\pdf\\',fpdf))
      print(fpdf+" - удален")

   excelfilelist=os.listdir(os.path.dirname(__file__)+'\\download\\')
   for fn_ in excelfilelist:
      if fn_.endswith(".xlsx"):
         EXCEL_FILE_PATH = os.path.dirname(__file__)+'\\download\\'+fn_
         DIR_PATH = os.path.dirname(__file__)

         if os.path.exists(EXCEL_FILE_PATH):
            test1=pyexcel.get_records(file_name=EXCEL_FILE_PATH, name_columns_by_row=0)
            df=pd.DataFrame(test1)[columns]

            data={}
            for row in df.values:
               n_cert = row[1]
               tr_date = datetime.date.isoformat(row[0])
               ts = row[2]
               url = row[3]
               if not (n_cert in data) and n_cert and tr_date and ts and url:
                  data[n_cert] = [tr_date, ts, url]
                  name = f'{DIR_PATH}/pdf/Серт_{n_cert}_дата_{tr_date}_тс_{ts}.pdf'
                  print(url)
                  try:
                     pdf_dowload(url, name)
                  except Exception:
                     try:
                        pdf_dowload(url, name)
                     except Exception:
                        try:
                           pdf_dowload(url, name)
                        except Exception:
                           continue
                     continue