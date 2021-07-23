# Dynatrace Host Units Report 

This script is able to extract the details of the HUs consumption for one or several specific Dynatrace tenants. 

## Description
- List of all monitored hosts on the tenant with the amount of HU consummed for each one (in separate Excel sheets)
- Amount of HU consumed per Management Zone 
- Excel file generation 

## Installation

##### Set file permissions
``` bash
cd ReportHU/
sudo chown u+x getHostUnits.pl
sudo chown u+w getHostUnits.ini
```

##### Set configuration file 
Edit the configuration file :
``` bash
$ sudo vi getHostUnits.ini
```
Modify one or more sections (depending on the number of tenant(s) managed) starting with "[tenant-XXXX]"
- Replace the XXXX value by the name of your Dynatrace tenant
- Specify the "tenant" with the value of the tenant ID (example in SaaS mode: tde91925)
- Specify the "token" value after generating a token with the correct rights 

``` 
$ cat getHostUnits.ini

[global]
logfile=getHostUnits.log
exportfileHU=dynatrace_report_HU.xlsx
resolution=M

[tenant-TenantName1]
tenant=<tenantid1>
token=<API-Token1>

[tenant-TenantName2]
tenant=<tenantid2>
token=<API-Token2>
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
$ grep "^use" getHostUnits.pl | egrep -v "constant|strict"
use LWP::UserAgent;
use JSON;
use JSON::XS qw(encode_json);
use Config::Tiny;
use DateTime;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
```

Depending on the execution environment, these modules are maybe already installed. If it's not the case, you have several options: 
1) with CPAN :
cpan -i <module_name>
2) with your package manager : 
- on Ubuntu/DEBIAN : apt install libexcel-writer-xlsx-perl
- on Centos/RHEL : yum install 
 
##### Command-line
 ```
 ./getHostUnits.pl
INFO [07/23/21 15:08:01] main - Start ./getHostUnits.pl
INFO [07/23/21 15:08:02] createExcel - Excel file creation
INFO [07/23/21 15:08:02] getHostList - https://tde95935.live.dynatrace.com/api/v1/entity/infrastructure/hosts?relativeTime=day
INFO [07/23/21 15:08:05] getHostList - Response : 200 OK
INFO [07/23/21 15:08:06] createExcel - Data writing
INFO [07/23/21 15:08:06] createExcel - Chart insertion
INFO [07/23/21 15:08:07] createExcel : Excel file "dynatrace_report_HU.xlsx" generated in the current directory
INFO [07/23/21 15:08:13] main - End of script
``` 
## Results

![alt text](https://github.com/mackermann2/dynatrace/blob/main/ReportHU/screenshot_exportHU_sample.png)
