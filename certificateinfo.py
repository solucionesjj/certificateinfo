import ssl
import socket
import html
import os
import openpyxl
import datetime as dt

# Si se presentan problemas con los certificados SSL por favor ejecutar:
# pip install –upgrade certifi
# o
# pip3 install –upgrade certifi
 
os.system("clear")
now = dt.datetime.now()
now = now.strftime("%Y%m%d_%H%M%S")


def getCertificateInfoFor(url):
    ipAddress = ""
    issuer = ""
    fromDate = ""
    toDate = ""
    subject = ""
    commonName = ""
    error = ""
    try:
        ipAddress = socket.gethostbyname(url)
    except Exception as e:
        ipAddress = str(e)

    try:
        context = ssl.create_default_context()
        with socket.create_connection((url, 443),timeout=10) as sock:
            with context.wrap_socket(sock, server_hostname=url) as ssock:
                cert = ssock.getpeercert()
                try:
                    issuer = dict(x[0] for x in cert['issuer']).get('organizationName','') +" - "+ dict(x[0] for x in cert['issuer']).get('commonName','')
                except Exception as e:
                    issuer =  cert['issuer']
                
                fromDate = cert['notBefore']
                toDate = cert['notAfter']

                try:
                    subject = dict(x[0] for x in cert['subject']).get('commonName','')
                except Exception as e:
                    subject = cert['subject']
                
                try:
                    commonName = dict(cert['subjectAltName']).get('DNS', '')
                except Exception as e:
                    commonName = cert['subjectAltName']
                
    except Exception as e:
        error = str(e)
    info = [url,ipAddress,issuer,fromDate,toDate,subject,commonName,error]
    print(info)
    return info

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

actualDomainCount = 1
totalDomains = len(domains)
for url in domains:
    print(f"Domain ({actualDomainCount}/{totalDomains}): {url}")
    info = getCertificateInfoFor(url)
    result = 'Ok' if info[6] == '' else info[6]
    info[6] = result
    print(f"Result: {result}")
    domainsInfo.append(info)
    actualDomainCount = actualDomainCount + 1
    print("")

wb = openpyxl.Workbook()
ws = wb.active
for data in domainsInfo:
    ws.append(data)
fileName = f"{now}_certificateInfo.xlsx"
wb.save(fileName)

print(f"File {fileName} generated successful.")

end = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
print(f"End: {end}")
