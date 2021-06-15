# Excel order Add-In
Generates an order table from tables exported from Altus Vario.

## Installation
1. Download at http://excel-order-addin.infinityfreeapp.com/download/.
2. Unzip archive.
3. Follow https://www.youtube.com/watch?v=alDwd8ghO7A&ab_channel=UriGorenUriGoren to install the certificate.
4. Install by double-clicking `setup.exe`.
5. Open Excel and use the plugin. It is added as **Order Add-In** in the ribbon. 
6. Go to `File -> Options -> Add-ins`. The *Name* is **ExcelOrderAddIn**. The *Location* is automatically scanned for updates. By placing newer version of add-in in the *Location*, the plugin will be updated. 
7. You can delete the installation folder if you want to, the plugin will still work.

## Other
### Certificate
It has been generated using the following commands. It will not expire and there is no need to generate it again.
1. Open PowerShell as Administrator.
2. `cd "C:\Program Files (x86)\Windows Kits\10\bin\10.0.16299.0\x64"`
3. `.\makecert.exe /n "CN=PetrJanikExcelOrderAddIn" /r /h 0 /eku "1.3.6.1.5.5.7.3.3,1.3.6.1.4.1.311.10.3.13" /e "01/16/2174" /sv PetrJanikExcelOrderAddIn.pvk PetrJanikExcelOrderAddIn.cer`
4. `pvk2pfx -pvk PetrJanikExcelOrderAddIn.pvk -spc PetrJanikExcelOrderAddIn.cer -pfx PetrJanikExcelOrderAddIn.pfx`
5. Copy the three generated files to project repository.
6. `open ExcelOrderAddIn project Properties -> Signing -> Select from File... -> select PetrJanikExcelOrderAddIn.pfx`
