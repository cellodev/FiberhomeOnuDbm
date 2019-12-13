This application makes a signal query (Dbm) on all authorized "ONUs" an Olt and exports it in xlsx format.

It is compatible with Olt brand Fiberhome firmware RP1000. For smooth operation requires Python3+,
telnetlib library, xlsxwriter library installed. 
'pip install telnetlib'
'pip install xlsxwriter'.

Usage:
Telnet_Olt 192.168.1.1 GEPON GEPON 23 (Default)
Telnet_Olt <IP> <User> <Password> <Port>
