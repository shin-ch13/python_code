import openpyxl
from bs4 import BeautifulSoup as bs
from urllib.request import *
from urllib.parse import *
from os import makedirs
import os.path, time, re, sys

API = "https://nvd.nist.gov/vuln/detail/"

# function to analize html
def analize_html(html,i,sheet):
  soup = bs(html,"html.parser")
  base_score = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(1) > a > span:nth-of-type(1)").string
  print(base_score)
  sheet['B'+str(i)] = base_score
  
  base_score_severity = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(1) > a > span:nth-of-type(2)").string
  print(base_score_severity)
  sheet['C'+str(i)] = base_score_severity
  
  cvssv2_vector = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(1) > span:nth-of-type(1)").string
  print(cvssv2_vector)
  sheet['D'+str(i)] = cvssv2_vector
  
  impact_subscore = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(1) > span:nth-of-type(2)").string
  print(impact_subscore)
  sheet['D'+str(i)] = impact_subscore
  
  exploitability_score = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(1) > span:nth-of-type(3)").string
  print(exploitability_score)
  sheet['E'+str(i)] = exploitability_score
  
  cvssv2_av = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(2) > span:nth-of-type(1)").string
  print(cvssv2_av)
  sheet['F'+str(i)] = cvssv2_av
  
  cvssv2_ac = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(2) > span:nth-of-type(2)").string
  print(cvssv2_ac)
  sheet['G'+str(i)] = cvssv2_ac
  
  cvssv2_au = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(2) > span:nth-of-type(3)").string
  print(cvssv2_au)
  sheet['H'+str(i)] = cvssv2_au
  
  cvssv2_c = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(2) > span:nth-of-type(4)").string
  print(cvssv2_c)
  sheet['I'+str(i)] = cvssv2_c
  
  cvssv2_i = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(2) > span:nth-of-type(5)").string
  print(cvssv2_i) 
  sheet['J'+str(i)] = cvssv2_i
  
  cvssv2_a = soup.select_one("#p_lt_WebPartZone1_zoneCenter_pageplaceholder_p_lt_WebPartZone1_zoneCenter_VulnerabilityDetail_VulnFormView_Vuln2CvssPanel > p:nth-of-type(2) > span:nth-of-type(6)").string
  print(cvssv2_a)
  sheet['K'+str(i)] = cvssv2_a

  return 

# function to request HTML
def request_html(key,i,sheet):
  # request url
  url = API + str(key)
  # request html
  try:
    html = urlopen(url).read()
    html = html.decode("utf-8")
    print("request=",url)
  except:
    print("Faile Request")
    return
  analize_html(html,i,sheet)

if __name__ == '__main__':
  i=2
  wb = openpyxl.load_workbook('CVE.xlsx')
  sheet = wb.get_sheet_by_name('Sheet2')
  for row_of_cell_objects in sheet['A2':'A27']:
    for cell_obj in row_of_cell_objects:
      print(i, cell_obj.coordinate, cell_obj.value)
      request_html(cell_obj.value,i,sheet)
    i=i+1
  wb.save('CVE.xlsx')
