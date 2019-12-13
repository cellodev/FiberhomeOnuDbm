import telnetlib
import re
import sys
import xlsxwriter

try:
    ipolt = "|192.168.1.1|"
    ipolt = str(sys.argv[1])
    if ipolt == "-h":
        print("This application makes a signal query (Dbm) on all authorized on an Olt and exports it in xlsx format.\nIt is compatible with Olt brand Fiberhome firmware RP1000. For smooth operation requires Python3+,\ntelnetlib library, xlsxwriter library installed. 'pip install telnetlib', 'pip install xlsxwriter'.")
        print("Usage:\nTelnet_Olt 192.168.1.1 GEPON GEPON 23 (Default)\nTelnet_Olt <IP> <User> <Password> <Port>")
        sys.exit()
except:
    if ipolt == "|192.168.1.1|":
        sys.exit("'Telnet_Olt -h' for help.")
    else:
        sys.exit("####Closed!####")
        
try:
    user = str(sys.argv[2])
except:
    user = "GEPON"
try:
    pwd = str(sys.argv[3])
except:
    pwd = "GEPON"
try:
    por = str(sys.argv[4])
except:
    por = 23

retorno = ""

#File
xbook = xlsxwriter.Workbook('My Query OLT.xlsx')
xsheet = xbook.add_worksheet('Query')
bold = xbook.add_format({'bold': True})
xsheet.write("A1", "SERIAL:", bold)
xsheet.write("B1", "DBM:", bold)
xsheet.write("C1", "SLOT:", bold)
xsheet.write("D1", "PON:", bold)

try:
    conn = telnetlib.Telnet(ipolt, por, timeout=10)
except:
    sys.exit("Error, 'Telnet_Olt -h' for help.")
conn.read_until(b"Login:", timeout = 2)
conn.write((user + "\n").encode('ascii'))
conn.read_until(b"Password:", timeout = 2)
conn.write((pwd + "\n").encode('ascii'))
c = conn.read_until(b"User>", timeout = 1)
if "Bad Password" in str(c):
    sys.exit("Bad User or Password")
conn.write(("EN\n").encode('ascii'))
conn.read_until(b"Password:", timeout = 2)
conn.write((pwd + "\n").encode('ascii'))
conn.read_until(b"#", timeout=1)
conn.write(("cd onu\n").encode('ascii'))
conn.read_until(b"onu", timeout=1)
conn.write(("show authorization slot all pon all\n").encode('ascii'))
conn.read_until(b"onu", timeout=1)
conn.write("cd ..\n".encode('ascii'))
conn.write("exit\n".encode('ascii'))
conn.write("exit\n".encode('ascii'))
saida = str(conn.read_all().decode('ascii'))
allonu = []
try:
    local = re.findall("FHTT(.*?)mac", saida)
    if len(local) == 0:
        local = re.findall("FHTT(.{15})", saida)
    for i in local:
        i = "FHTT"+i
        i = i.replace(" ", "")
        allonu.append(i)
except:
    print("an error occurred in locating the Onu")
x = 0
for seria in allonu:
    x = x+1
    conn = telnetlib.Telnet(ipolt, por, timeout=10)
    conn.read_until(b"Login:", timeout = 2)
    conn.write((user + "\n").encode('ascii'))
    conn.read_until(b"Password:", timeout = 2)
    conn.write((pwd + "\n").encode('ascii'))
    conn.read_until(b"User>", timeout = 2)
    conn.write(("EN\n").encode('ascii'))
    conn.read_until(b"Password:", timeout = 2)
    conn.write((pwd + "\n").encode('ascii'))
    conn.read_until(b"#", timeout=1)
    conn.write(("cd onu\n").encode('ascii'))
    conn.read_until(b"onu", timeout=1)
    conn.write(("show onu-authinfo phy-id {}\n").format(seria).encode('ascii'))
    conn.read_until(b"onu", timeout=1)
    conn.write("cd ..\n".encode('ascii'))
    conn.write("exit\n".encode('ascii'))
    conn.write("exit\n".encode('ascii'))
    saida = str(conn.read_all().decode('ascii'))
    local = []
    try:
        local = saida.split("ONU:",1)[1]
        local = local.split('OnuType')[0]
        local = local.split('-')
    except:
        local.append("Error collect")
    local1 = []
    for s in local:
        s = s.replace(" ", "")
        s = s.strip()
        local1.append(s)
    model = saida.split("OnuType",1)[1]
    model = model.split("Phy-id")[0]
    model = model.split(":",1)[1]
    model = model.split("(")[0]
    model = model.replace(" ", "")
    
    #Find Dbm
    conn = telnetlib.Telnet(ipolt, por, timeout=10)
    conn.read_until(b"Login:", timeout = 2)
    conn.write((user + "\n").encode('ascii'))
    conn.read_until(b"Password:", timeout = 2)
    conn.write((pwd + "\n").encode('ascii'))
    conn.read_until(b"User>", timeout = 1)
    conn.write(("EN\n").encode('ascii'))
    conn.read_until(b"Password:", timeout = 2)
    conn.write((pwd + "\n").encode('ascii'))
    conn.read_until(b"#", timeout=1)
    conn.write(("cd onu\n").encode('ascii'))
    conn.read_until(b"onu", timeout=1)
    conn.write(("show optic_module slot {} pon {} onu {}\n").format(local1[0], local1[1], local1[2]).encode('ascii'))
    conn.read_until(b"onu", timeout=1)
    conn.write("cd ..\n".encode('ascii'))
    conn.write("exit\n".encode('ascii'))
    conn.write("exit\n".encode('ascii'))
    saida = str(conn.read_all().decode('ascii'))
    try:
        sinal = saida.split("RECV POWER",1)[1]
        sinal = sinal.split("(")[0]
        sinal = sinal.split(":",1)[1]
        sinal = sinal.replace(" ", "")
    except:
        retorno = "Error"
        sinal = "Unknown"
    print(str(x)+ " of " +str(len(allonu))+ " || ONU: "+seria+ " || Dbm: " +sinal)

    p = []
    p.append(seria)
    p.append(sinal)
    p.append(local1[0])
    p.append(local1[1])
    xsheet.write_row(x, 0, p)
xbook.close()
