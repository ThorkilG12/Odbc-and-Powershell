# Odbc-and-Powershell
It is quite easy to use ODBC with Powershell.  
The tricky part is living in Europe where more letters than A-Z exist.  
Notice that every file has encoding. No matter what.  
I'm using Powershell 5.1 on a Windows 2019 Server. The server displays tekst in US, but has da-DK as locale language
I have downloaded and installed the US version of `AccessDatabaseEngine_X64.exe`  
Let's create a setup:  
<img width="430" alt="image" src="https://user-images.githubusercontent.com/12120277/209286073-00175c2f-e9f2-47c0-8ae7-71fd0d3f07f7.png">  
I have named the Data Source and pointed to the folder:  
<img width="332" alt="image" src="https://user-images.githubusercontent.com/12120277/209288193-3f62797c-ecb8-45de-bf77-a3a56934803e.png">  
I have this file - encoded in UTF-8 - in the folder `D:\deledata`. Filename is `import.txt`  
<img width="123" alt="image" src="https://user-images.githubusercontent.com/12120277/209286657-d5e3fbb1-677a-483f-8e1b-2ea2c5c7a0e0.png">  
```PS
$query='select * from import.txt'
$conn = New-Object System.Data.Odbc.OdbcConnection
$conn.ConnectionString = "DSN=myText;"
$conn.open()
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$ds = New-Object system.Data.DataSet
$numrows = (New-Object system.Data.odbc.odbcDataAdapter($cmd)).fill($ds) 
$conn.close()
Write-host "Number of rows returned: $numrows"
$ds.tables.rows
```
Result is this:  
<img width="153" alt="image" src="https://user-images.githubusercontent.com/12120277/209288685-75f0c0d5-5173-48be-9912-86108c2ba1a0.png">  
The driver has used first line as heading, but the semi-colon as delimiter is not used. More config is needed.  
<img width="367" alt="image" src="https://user-images.githubusercontent.com/12120277/209290823-a6f00beb-3027-4aa6-9bc0-95577e2623f8.png">  
After this configuration, notice that a file called `schema.ini` has been created in `d:\deledata`  
with this content:  
```
[import.txt]
ColNameHeader=True
Format=Delimited(;)
MaxScanRows=25
CharacterSet=OEM
```
Now the result is like below, and all english talking people are happy.    
<img width="147" alt="image" src="https://user-images.githubusercontent.com/12120277/209291411-fd480e83-43b1-4ce8-874b-6b59ea7b8f99.png">  
But not the people from EU... Lets add the boy Åbjørn from the polish town Łódź. I always use this town, since it has letters that fall outside Western Codepage.   Western was the solution back in around 1980, but not anymore. Now we have this:    
<img width="115" alt="image" src="https://user-images.githubusercontent.com/12120277/209292297-26bf083e-3061-4c59-9aaf-0374ec573a14.png">  
Clearly not ok. The driver allow us to choose between ANSI and OEM, but that is not good enough. We need to change one line in the .ini file:  
`CharacterSet=65001`  
Now even the people from EU are happy !  
<img width="144" alt="image" src="https://user-images.githubusercontent.com/12120277/209292967-f7774ca5-934e-413c-90db-e253be76061f.png">  

  --oOo--
