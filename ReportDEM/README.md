# Dynatrace Digital Experience Monitoring Report 

This script is able to extract the details of the DEM consumption for one or several specific Dynatrace tenants. 

## Description
- Amount of DEM consumed on the whole tenant and per Management Zone 
- Breakdown of the DEM licenses consumption via Excel comments (tooltips) : RUM (Real User Monitoring) or Synthetics (HTTP Monitor/Browser Monitor)
- Following of the percentage of used licenses in the global pool (since DEM licenses are not distributed by tenant)
- Excel file generation 

## Installation

##### Set file permissions
``` bash
cd ReportDEM/
sudo chown u+x getDEMLicenses.pl
sudo chown u+w getDEMLicenses.ini
```

##### Set configuration file 
Edit the configuration file :
``` bash
$ sudo vi getDEMLicenses.ini
```
Modify one or more sections (depending on the number of tenant(s) managed) starting with "[tenant-XXXX]"
- Replace the XXXX value by the name of your Dynatrace tenant
- Specify the "tenant" with the value of the tenant ID (example in SaaS mode: tde91925)
- Specify the "token" value after generating a token with the correct rights 
- Set a maximum limit of DEM for the tenant (example: 1 million)
- To get the DEM consumption of a specific perimeter, you can fill the following parameters ManagementZone1 and limitDEM_ManagementZone1 to filter on a specific Management Zone

``` 
$ grep -A8 "tenant-" ReportDEM/getDEMLicenses.ini
[tenant-XXXX]
tenant=<tenantid1>
token=<API-Token1>
limitDEM=1000000
ManagementZone1=DEV
limitDEM_ManagementZone1=100000
ManagementZone2=PROD
limitDEM_ManagementZone2=250000
```

##### Pre-requisites 
- API key must be generated on the Dynatrace tenant for the API v1 and with the following permission : "Access problem and event feed, metrics, and topology"

## Execution
##### Perl environment
- On MS Windows, you will need a perl environment. [Strawberry Perl](https://strawberryperl.com) could do the job.
- You can also use WSL (Windows Subsystem for Linux) if the feature is enabled on your operating system. 
- On Linux, Perl is often already installed with the system. So, a simple shell will be enough.
  
Note: some extra PERL modules are used (especially for the Excel file generation) and must be installed :

``` bash
$ grep "^use" ReportDEM/getDEMLicenses.pl | egrep -v "constant|strict"
use LWP::UserAgent;
use Data::Dumper;
use JSON;
use JSON::XS qw(encode_json);
use List::Util qw(sum);
use Config::Tiny;
use feature qw{say};
use DateTime;
use Date::Calc qw(:all);
use URI::Encode;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Spreadsheet::XLSX;
use Spreadsheet::ParseExcel::Utility qw(ExcelFmt);
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Text::Iconv;
```

Depending on the execution environment, these modules are maybe already installed. If it's not the case, you have several options: 
1) with CPAN :
cpan -i <module_name>
2) with your package manager : 
- on Ubuntu/DEBIAN : apt install libexcel-writer-xlsx-perl
- on Centos/RHEL : yum install 
 
##### Command-line
 ```
./getDEMLicenses.pl 
INFO [2021-09-09 15:05:30] main - Start ./getDEMLicenses.pl
INFO [2021-09-09 15:05:30] makeDate - Last month = 2021/08
INFO [2021-09-09 15:05:30] makeDateYear - List of the months in the past before the current date
INFO [2021-09-09 15:05:30] makeDateYear - Loop on all the months 
INFO [2021-09-09 15:05:30] getOption - Resolution: M
INFO [2021-09-09 15:05:30] checkExist - File detection dynatrace_report_DEM.xlsx : OK.
INFO [2021-09-09 15:05:30] main - Count of the licenses of the previous month
INFO [2021-09-09 15:05:31] checkLastMonth - $lastMonth = 07/2021
INFO [2021-09-09 15:05:31] duplicateExcel - Clone of the worksheet "Global".
INFO [2021-09-09 15:05:31] duplicateExcel - Clone of the worksheet "tenant-MyTenant1".
INFO [2021-09-09 15:05:31] updateExcel - tenant-MyTenant1
INFO [2021-09-09 15:05:31] updateExcel - tenant-MyTenant1 > Management Zone DEV
INFO [2021-09-09 15:05:34] updateExcel - tenant-MyTenant1 > Management Zone PROD
INFO [2021-09-09 15:05:38] duplicateExcel - Clone of the worksheet "tenant-MyTenant2".
INFO [2021-09-09 15:05:38] makeChart - Build of the chart on the first sheet
``` 
## Results

![alt text](https://github.com/mackermann2/dynatrace/blob/main/ReportDEM/screenshot1_exportDEM_sample.png)
![alt text](https://github.com/mackermann2/dynatrace/blob/main/ReportDEM/screenshot2_exportDEM_sample.png)
![alt text](https://github.com/mackermann2/dynatrace/blob/main/ReportDEM/screenshot3_exportDEM_sample.png)