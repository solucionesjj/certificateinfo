import ssl
import socket
import html
import os
import openpyxl
import datetime as dt
 
os.system("clear")
now = dt.datetime.now()
now = now.strftime("%Y%m%d_%H%M%S")


def getCertificateInfoFor(url):
    issuer = ""
    fromDate = ""
    toDate = ""
    subject = ""
    commonName = ""
    error = ""
    try:
        context = ssl.create_default_context()
        with socket.create_connection((url, 443),timeout=5) as sock:
            with context.wrap_socket(sock, server_hostname=url) as ssock:
                cert = ssock.getpeercert()
                issuer = dict(x[0] for x in cert['issuer']).get('organizationName','') +" - "+ dict(x[0] for x in cert['issuer']).get('commonName','')
                fromDate = cert['notBefore']
                toDate = cert['notAfter']
                subject = dict(x[0] for x in cert['subject']).get('commonName','')
                commonName = dict(cert['subjectAltName']).get('DNS', '')
    except Exception as e:
        error = str(e)
    # print([url,issuer,fromDate,toDate,subject,commonName,error])
    return [url,issuer,fromDate,toDate,subject,commonName,error]

domains = []

print(f"Introduzca los dominios uno por línea para encontrar la información SSL, para terminar coloque !q: ")
while True:
    domain = html.escape(input())
    if domain == "!q":
        break
    domains.append(domain)

os.system("clear")
print(f"Start: {now}")

domains = sorted(list(set(domains)))

domainsInfo = []

for url in domains:
    print(f"Domain: {url}")
    info = getCertificateInfoFor(url)
    result = 'Ok' if info[6] == '' else info[6]
    info[6] = result
    print(f"Result: {result}")
    domainsInfo.append(info)

wb = openpyxl.Workbook()
ws = wb.active
for data in domainsInfo:
    ws.append(data)
fileName = f"{now}_certificateInfo.xlsx"
wb.save(fileName)

print(f"File {fileName} generated successful.")

end = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
print(f"End: {end}")
