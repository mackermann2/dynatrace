#!/bin/perl -w
#--------------------------------------------------------------------------------------------------------------------------------------------------------------#
# AUTHOR        : Ackermann Matthieu
# DATE          : 2020-10-01
# DESCRIPTION   : Export of the consumption of the Dynatrace DEM licences in an Excel file (from Real User Monitoring and Synthetic Monitoring) 
# UPDATE        : 2020-10-27 : Only update the last line of the excel file (corresponding to the previous month) in case of the file already exists. 
#				  Very useful when the limit of API usage has been reached (HTTP response code : 429 - Too many requests)
#				  > Limitations: all the previous comments (details consumption for each month) are lost because the used PERL modules don't support the reading 
#				  of the existant comments (and Win32::OLE which could be the solution works only in a Windows environment (not WSL or Cygwin))
#				  2021-01-05 : Problem of script execution in January (corrections made in makeDateYear() function)
#				  2021-09-09 : Adding the evolution of the DEM consumption between each month in a separate column 
#--------------------------------------------------------------------------------------------------------------------------------------------------------------#
use strict;
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

#EXCEL MODULES
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Spreadsheet::XLSX;
use Spreadsheet::ParseExcel::Utility qw(ExcelFmt);
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Text::Iconv;

#CONSTANTS
use constant L_ERROR 	=> "ERROR";
use constant L_WARNING	=> "WARNING";
use constant L_INFO	=> "INFO";
use constant SUCCESS	=> 0;
use constant ERROR      => 1;
use constant WARNING    => 2;


#---- MAIN --------------------------------------------------------------------#
# LOADING OF CONFIGURATION FILE
my $conf = Config::Tiny->new;
$conf = Config::Tiny->read( "getDEMLicenses.ini" );

# INITIALIZATION LOG FILE
my $LOG_FILE = $conf->{"global"}->{"logfile"};

&log( L_INFO, "main - Start $0");
# INITIALIZATION OPTIONS
my $date_start;
my $date_end;
my $dt = DateTime->now();
my $resolution;

#VARIABLES
my $exportFile = $conf->{"global"}->{"exportfileDEM"};
my $tenant;
my $apitoken;
my $url;
my $contractRenewalDate = $conf->{"global"}->{"contractRenewalDate"};

my $conversion;
my %appList;
my %syntheticList;
my %httpMonitorList;
my $totalDEM=0;
my %hDate   = ();
my @hDateYear = ();
my %hStat   = ();

#DATES MANAGEMENT
&makeDate(\%hDate);
&makeDateYear(\@hDateYear);
getOption();

#DECLARATION OF LWP USER AGENT
my $ua = LWP::UserAgent->new;
$ua->timeout(100);

#CHECK IF THE EXPORT FILE EXISTS
my $existFile   = &checkExist($exportFile);

#IF YES : LOADING OF THE MONTH N-1
if ($existFile == 1){
    &log( L_INFO, "main - Count of the licenses of the previous month" );

	#READING OF EXCEL FILE
	my $converter   = Text::Iconv->new("utf-8", "windows-1251");
	my $excel_r     = Spreadsheet::XLSX->new($exportFile, $converter);

    # CHECK THE DATE OF THE LAST EXPORT
	if ( &checkLastMonth($excel_r) eq $hDate{"last_month"}."/".$hDate{"last_year"} ) {
		&log( L_ERROR, "main - Export already generated for the previous month" );
		&log( L_ERROR, "------" );
		exit 1;
    }
	# LOADING OF THE EXCEL FILE TO MODIFY (Spreadsheet::XLSX NOT ABLE TO WRITE)
	my $excel_w     = Excel::Writer::XLSX->new(substr($exportFile,0,length($exportFile)-5)."_".$hDate{"last_year"}.$hDate{"last_month"}.".xlsx");
	    	
	#LOOP ON ALL THE SHEETS
	my $iSheet=0;
    foreach my $sheet (@{$excel_r -> {Worksheet}}) {
			$iSheet++;
		    #my $worksheet_title = $sheet->{Name};
	        #&log( L_INFO, "main - Sheet $worksheet_title" );

			# CREATION OF THE SHEET
    		my $worksheet = $excel_w->add_worksheet($sheet->get_name());

			# EXCEL FILE REPLICATION
    		&duplicateExcel($excel_w,$sheet,$worksheet);
			
			# UPDATE WITH THE DATA OF THE LAST MONTH
			&updateExcel($excel_w,$sheet,$worksheet,\%hDate);
			
			# CHART CREATION AFTER THE BUILD OF THE LAST SHEET
			if ($iSheet eq $excel_r->{SheetCount}){
				&makeChart($excel_w,$sheet,$worksheet,\%hDate);
			}
	}
	$excel_w->close();
}
else{
	# LOADING OF THE EXCEL FILE TO MODIFY (Spreadsheet::XLSX NOT ABLE TO WRITE)
	my $excel_w     = Excel::Writer::XLSX->new(substr($exportFile,0,length($exportFile)-5)."_".$hDate{"last_year"}.$hDate{"last_month"}.".xlsx");
	my $sheet = $excel_w->add_worksheet("Global");
	# CREATE THE FILE FROM SCRATCH
	&createExcel($excel_w,$sheet);
	$excel_w->close();
}

#---- log ---------------------------------------------------------------------#
#
# WRITING OF THE LOGS
#
# Entree :
#	logType = Type de log : L_ERROR, L_INFO, L_WARNING
#	logMess = Message
# Return :

sub log (){

	my $logType = $_[0];
	my $logMess = $_[1];
	my $dt = DateTime->now;
    my $ymd = $dt->ymd;
    my $hms = $dt->hms;
    my $Date = "[".$ymd." ".$hms."]";
	#chop (my $Date = `/bin/date "+[%x %H:%M:%S]"`);

	open LOGFILE, ">> $LOG_FILE";
	print LOGFILE "$logType $Date $logMess\n";
	close LOGFILE;

	print "$logType $Date $logMess\n";

	return;
}

# FUNCTION usage()
sub usage () {
	print "Usage: $0 [--help] \n
	List of parameters:
	--help: To display this help
	--resolution: The desired resolution. Possible value (default:M):  m(minutes), h(hours), d(days), w(weeks), M(months), y(years)
	--from : The start of the requested timeframe. If not set, the first day of the last month is used.
	--to : The end of the requested timeframe. If not set, the last day of the last month is used.

	EXEMPLE:
	perl $0 --resolution M --from 20200801 --to 20200831\n
	\n";
	exit 0;
}
# GETTING OPTIONS
sub getOption {
	use Getopt::Long;
	my $help;
	GetOptions ("help" => \$help,
		"resolution=s" => \$resolution,
		"from=s" => \$date_start,
		"to=s" => \$date_end)
		or die("Error in command line arguments\n");

	if (defined $help) { 
		&usage(); 
	}
	unless ( defined $resolution ) { 
		$resolution="M"; 
	}
	#unless ( defined $date_start ) { 
		#$date_start=DateTime->new( year => $dt->year, month => $dt->month-1, day => 1, hour=>0, minute=>0,second=>0,time_zone=>'Europe/Paris' );   
		#$date_start=DateTime->new( year => 2020, month => 12, day => 1, hour=>0, minute=>0,second=>0,time_zone=>'Europe/Paris' );   
	#}
	#unless ( defined $date_end ) { 
		#$date_end=DateTime->last_day_of_month( year => $dt->year, month => $dt->month-1, hour=>23, minute=>59,second=>59,time_zone=>'Europe/Paris'  );  
		#$date_end=DateTime->last_day_of_month( year => 2020, month => 12, hour=>23, minute=>59,second=>59,time_zone=>'Europe/Paris'  );  
	#}
	&log( L_INFO, "getOption - Resolution: $resolution");
	#&log( L_INFO, "getOption - Dates range : from $date_start to $date_end");
}

sub execCommand{
	my $cmd  = $_[0];
	my $resultat =`$cmd`;
	chomp($resultat);
	return $resultat;
}

sub getAppList{
	my $tenant = $_[0];
	my $token = $_[1];
    my $date_start = $_[2];
	my $date_end = $_[3];

	# WEB APPLICATION
	my $url="https://$tenant.live.dynatrace.com/api/v2/entities?entitySelector=type(%22APPLICATION%22)&from=$date_start&to=$date_end";
	my $response = $ua->get($url, "Content-Type" => "application/json", "Authorization" => "Api-Token $token");
	my $DT_JSON=$response->decoded_content;
	my $decoded_json=decode_json($DT_JSON);
	my $jsonarray = $decoded_json->{entities};
		foreach my $item (@$jsonarray) {
			$appList{$item->{'entityId'}} = $item->{'displayName'};
		}
	# MOBILE APPLICATION
	$url="https://$tenant.live.dynatrace.com/api/v2/entities?entitySelector=type(%22MOBILE_APPLICATION%22)&from=$date_start&to=$date_end";
	$response = $ua->get($url, "Content-Type" => "application/json", "Authorization" => "Api-Token $token");
	$DT_JSON=$response->decoded_content;
	$decoded_json=decode_json($DT_JSON);
	$jsonarray = $decoded_json->{entities};
	foreach my $item (@$jsonarray) {
		$appList{$item->{'entityId'}} = $item->{'displayName'};
	}
	# CUSTOM APPLICATION
	$url="https://$tenant.live.dynatrace.com/api/v2/entities?entitySelector=type(%22CUSTOM_APPLICATION%22)&from=$date_start&to=$date_end";
	$response = $ua->get($url, "Content-Type" => "application/json", "Authorization" => "Api-Token $token");
	$DT_JSON=$response->decoded_content;
	$decoded_json=decode_json($DT_JSON);
	$jsonarray = $decoded_json->{entities};
	foreach my $item (@$jsonarray) {
		$appList{$item->{'entityId'}} = $item->{'displayName'};
	}
	#print Dumper %appList;
	#print Dumper $jsonarray;
}

sub getAppName{
	my $entityId = $_[0];
	return $appList{"$entityId"};
}

sub getSyntheticList{
	my $tenant = $_[0];
	my $token = $_[1];
	my $url="https://$tenant.live.dynatrace.com/api/v1/synthetic/monitors";
	my $response = $ua->get($url, "Content-Type" => "application/json", "Authorization" => "Api-Token $token");
	my $DT_JSON=$response->decoded_content;
	#my $decoded=decode_json($DT_JSON);
	my $decoded_json = decode_json($DT_JSON);
	my $jsonarray = $decoded_json->{monitors};
	foreach my $item (@$jsonarray) {
		$syntheticList{$item->{'entityId'}} = $item->{'name'};
	}
	#print Dumper %syntheticList;
	#print Dumper $jsonarray;
}

sub getSyntheticName{
	my $entityId = $_[0];
	return $syntheticList{"$entityId"};
}

sub getHttpMonitorList{
	my $tenant = $_[0];
	my $token = $_[1];
    my $date_start = $_[2];
	my $date_end = $_[3];

	my $url="https://$tenant.live.dynatrace.com/api/v2/entities?entitySelector=type(%22HTTP_CHECK%22)&from=$date_start&to=$date_end";
	my $response = $ua->get($url, "Content-Type" => "application/json", "Authorization" => "Api-Token $token");
	my $DT_JSON=$response->decoded_content;
	my $decoded_json=decode_json($DT_JSON);
	my $jsonarray = $decoded_json->{entities};

		foreach my $item (@$jsonarray) {
			$httpMonitorList{$item->{'entityId'}} = $item->{'displayName'};
		}
	#print Dumper %httpMonitorList;
	#print Dumper $jsonarray;
}

sub getHttpMonitorName{
	my $entityId = $_[0];
	return $httpMonitorList{"$entityId"};
}

#---- makeDateYear ------------------------------------------------------------#
#
# Dates Generation
#
# Entree :
#       \@hDateYear = Reference on an empty table of hash
# Sortie :
#       \@hDateYear = Hash updated
#           start   = 1st day of the month to export
#           end     = last day of the month to export
# Return :
#
sub makeDateYear {
	&log( L_INFO, "makeDateYear - List of the months in the past before the current date" );
    my $refHdate = $_[0];
    DateTime->DefaultLocale('fr_FR');
    my $nb_month = DateTime->now->subtract(months => 1)->month;
		&log( L_INFO, "makeDateYear - Loop on all the months " );
        for ( my $i = 0; $i < $nb_month; $i++ ) {
            my $subMonth = $nb_month - $i;
            my $dt = DateTime->now->subtract(months => $subMonth);
            my $last_year       = $dt->year;
            my $last_month      = $dt->strftime("%m");
            my $number_of_days  = Days_in_Month($last_year,$last_month);

            $refHdate->[$i]->{"start"}      = "$last_year-$last_month-01T00:00:00";
            $refHdate->[$i]->{"end"}        = "$last_year-$last_month-$number_of_days"."T23:59:59";
            $refHdate->[$i]->{"last_year"}  = "$last_year";
            $refHdate->[$i]->{"last_month"} = "$last_month";
			#&log( L_INFO, "makeDateYear - $last_year/$last_month" );
        }
    return;
}

#---- makeDate ----------------------------------------------------------------#
#
# Dates Generation
#
# Entree :
#       \@hDateYear = Reference on an empty table of hash
# Sortie :
#       \@hDateYear = Hash updated
#           start   = 1st day of the month to export
#           end     = last day of the month to export
# Return :
#
sub makeDate {
	
    my $refHdate = $_[0];

    DateTime->DefaultLocale('fr_FR');
    my $dt = DateTime->now->subtract(months => 1);
    my $last_year    = $dt->year;
    my $last_month   = $dt->strftime("%m");
    my $number_of_days  = Days_in_Month($last_year,$last_month);
    my $gmt="GMT+01:00";

    $refHdate->{"start"}         = "01/$last_month/$last_year+00:00+$gmt";
    $refHdate->{"end"}           = "$number_of_days/$last_month/$last_year+23:59+$gmt";
    $refHdate->{"last_year"}	 = "$last_year";
    $refHdate->{"last_month"}    = "$last_month";

	&log( L_INFO, "makeDate - Last month = $last_year/$last_month" );
    return;

}

#---- checkExist -----------------------------------------------------------#
#
# Check if the Excel file exists in the destination directory
#
# Entree :
#
# Sortie :
#
# Return : 1 (file exists) or 0 (file does not exist)
#
sub checkExist {
	my $exportFile	=  $_[0];
  	if (-e $exportFile){
        &log( L_INFO, "checkExist - File detection $exportFile : OK." );
		return 1;
    }
    else
    {
		&log( L_INFO, "checkExist - File detection $exportFile : NOK." );
		return 0;
    }
}

#---- checkLastMonth ----------------------------------------------------------#
#
# Get the month to export
#
# Entree :
#
# Sortie :
#
# Return :
#       lastMonth = last month to export
sub checkLastMonth {
    my $refExcel = $_[0];
    my $lastMonth;

    # checking of the second sheet (the first is the chart)
    my $sheet = @{$refExcel -> {Worksheet}}[1];
	my $last_row=($sheet->{MaxRow})-3;
    #&log( L_INFO, "checkLastMonth - \$MaxRow = ". $last_row );

#    my ( $col_min, $col_max ) = $sheet->col_range();
#    my $cell = $sheet->get_cell(1, 0);
#    defined $sheet->{MaxRow} && $iR <= $sheet->{MaxRow} ; $iR++
    my $cell = $sheet -> {Cells} [$last_row] [0];

	if ($cell){
		my $cell = $sheet -> {Cells} [$last_row] [0];
	}
	else{
		my $last_line = $last_row;

		while (! $cell ){
			$last_line -= 1;
			$cell = $sheet -> {Cells} [$last_line][0];
		}
	}
    $lastMonth = ExcelFmt('mm/yyyy', $cell -> {Val});
    &log( L_INFO, "checkLastMonth - \$lastMonth = $lastMonth" );
    return $lastMonth;
}

#---- paramSheet --------------------------------------------------------------#
#
# Setting of each sheet in the Excel file
#
# Entree :
#       refSheet = object Excel sheet
#
# Sortie :
#
# Return :
#

sub paramSheet {
    my ($refSheet) = $_[0];

	# Size of the columns
	$refSheet->set_column(0,0,22);
    $refSheet->set_column(1,1,25);
	$refSheet->set_column(2,2,10);
	# Zoom to 85% (best view)
	# $refSheet->set_zoom(85);
}

#---- paramCell --------------------------------------------------------------#
#
# Format of the Excel cell
#
# Entree :
#	refExcel = object Excel file
# 	type = type of the cell
#
# Sortie :
#
# Return :
# 	format = format of the cell


sub paramCell {
    my ($refExcel,$type) = @_;
    my $format   = "";

	if ( $type eq 'entete' ){
		my $bg_color = $refExcel->set_custom_color(40, 196, 215, 155);
		$format = $refExcel->add_format(
		        font => 'Calibri',
		        center_across => 1,
		        bold => 1,
		        size => 11,
		        border => 1,
		        color => 'black',
		        bg_color => $bg_color,
		        border_color => 'black',
		        align => 'vcenter',
		        );
	}
	elsif ( $type eq 'mois' ){
		my $bg_color = $refExcel->set_custom_color(41, 218, 238, 243);
		$format = $refExcel->add_format(
		        font => 'Calibri',
		        center_across => 1,
		        bold => 1,
		        size => 11,
		        border => 1,
		        color => 'black',
		        bg_color => $bg_color,
		        border_color => 'black',
		        align => 'vcenter',
		        num_format => 'mmm-yy'
	        );
	}
	elsif ( $type eq 'total' ){
	my $bg_color = $refExcel->set_custom_color(42, 255, 195, 195);
	$format = $refExcel->add_format(
	        font => 'Calibri',
	        center_across => 1,
	        bold => 1,
	        size => 11,
	        border => 1,
	        color => 'red',
	        bg_color => $bg_color,
	        border_color => 'black',
	        align => 'vcenter'
        );
	}
	elsif ( $type eq 'body'){
		$format = $refExcel->add_format(
           font => 'Calibri',
           center_across => 1,
           bold => 0,
           size => 11,
           border => 1,
           color => 'black',
           bg_color => 'white',
           border_color => 'black',
           align => 'vcenter',
           );	
	}
	elsif ( $type eq 'trend'){
		$format = $refExcel->add_format(  
			font => 'Calibri',
  			bold => 1,
  			size => 11,
  			border => 1,
  			bg_color => 'white',
  			border_color => 'black',
  			align => 'right',
			num_format => 0x0a
  			);
	}
	else{
        $format = $refExcel->add_format(
            font => 'Calibri',
            center_across => 1,
            bold => 0,
            size => 11,
            border => 1,
            color => 'black',
            bg_color => 'white',
            border_color => 'black',
            align => 'right',
            );
        }

#	&log( L_INFO, "paramCell - \$format = |$format|" );
	return $format;
}

#---- getURL ---------------------------------------------------------------#
#
# URL Generation
#
# Entree :
#       url     = Access URL 
#       reponse = reference to the response of the function get
#       ua      = User Agent
# Sortie :
#       reponse = content updated
# Return :
#       SUCCESS = Request OK
#       ERROR   = Request KO
sub getURL {
	my $tenant = $_[0];
	my $token = $_[1];
    my $date_start = $_[2];
	my $date_end = $_[3];
	my $mgtZoneName = $_[4];
	my $refComment = $_[5];

	my $countLicence=0;
	#&log( L_INFO, "getURL - FROM $date_start TO $date_end");
	foreach my $metric (keys %{$conf->{"metrics"}}){
		#&log( L_INFO, "getURL - Metrics : ".$conf->{"metrics"}->{$metric});
		$url = "https://$tenant.live.dynatrace.com/api/v2/metrics/query?metricSelector=";
		$url.=$conf->{"metrics"}->{$metric};
		$url.="&resolution=$resolution&from=$date_start&to=$date_end";
		# IF FILTER ON A SPECIFIC MANAGEMENT ZONE 
		if (defined $mgtZoneName && $mgtZoneName ne ""){ 
			use URI::Encode qw(uri_encode uri_decode);
			my $encoded = uri_encode($mgtZoneName);
			my $dimension;

			if ($metric =~ "WebApplication"){
				$dimension="APPLICATION";
			}
			elsif($metric =~ "CustomApplication"){
				$dimension="CUSTOM_APPLICATION";
			}
			elsif($metric =~ "MobileApplication"){
				$dimension="MOBILE_APPLICATION";
			}
			elsif($metric =~ "HttpMonitor"){
				$dimension="HTTP_CHECK";
			}
			elsif($metric =~ "BrowserMonitor"){
				$dimension="SYNTHETIC_TEST";
			}
			elsif($metric =~ "ThirdPartyResult"){
				$dimension="EXTERNAL_SYNTHETIC_TEST";
			}
			$url.="&entitySelector=type($dimension),mzName($encoded)";
		}
		$conversion = $conf->{"conversion"}->{$metric};
		
		my $response = $ua->get($url, "Content-Type" => "application/json", "Authorization" => "Api-Token $token");
		my $json_data;
		if ( $response->is_success ) {
			$json_data = decode_json($response->decoded_content);	
		}
		else {
			&log( L_ERROR, "getURL - Requete en erreur : ".$response->status_line );
			&log( L_ERROR, "$url");
			return ERROR;
		}
		my $totalCount=$json_data->{'totalCount'};

		for (my $i = 0; $i < $totalCount; $i++) {
			my $dimension=$json_data->{'result'}->[0]->{'data'}->[$i]->{'dimensions'}[0];
			my $value=$json_data->{'result'}->[0]->{'data'}->[$i]->{'values'}[0];
			
			#if (defined $value){ $countLicence+=$value;}
			if (defined $dimension && $value ne 0){ 
				if ($dimension =~ "APPLICATION"){
					#print getAppName($dimension)." : $value\n";
					#print $refComment;
					$countLicence+=$value*$conversion;
					$$refComment.=getAppName($dimension)." : $value (user sessions) = ".($value*$conversion)." DEM\n";
				}
				elsif ($dimension =~ "SYNTHETIC"){
					#print "$dimension\n";
					#print getSyntheticName($dimension)." : $value\n";
					$countLicence+=$value*$conversion;
					$$refComment.=getSyntheticName($dimension)." : $value (synthetic actions) = ".($value*$conversion)." DEM\n";
				}
				elsif ($dimension =~ "HTTP_CHECK"){
					#print getHttpMonitorName($dimension)." : $value\n";
					$countLicence+=$value*$conversion;
					$$refComment.=getHttpMonitorName($dimension)." : $value (http check) = ".($value*$conversion)." DEM\n";
				}
			}
		}
}
#&log( L_INFO, "getURL - TOTAL DES LICENCES DEM CONSOMMES :".$countLicence );
return $countLicence;
}

#---- createExcel -----------------------------------------------------------#
#
# Creation o the file from scratch if it doesn't already exist (case of the 1st export)
#
# Entree :
#
# Sortie :
#
# Return :
#

sub createExcel {
    my $refExcel        = $_[0];
	my $refSheet        = $_[1];
    my $refWorkSheet    = $_[2];

            #FILE CREATION
			&log( L_INFO, "createExcel - Excel file creation (first export)" );
			my $i;
			foreach my $section (sort (keys %{$conf})){
				if ($section =~ /^tenant-/ ){
					my %mgtZone;
					&log( L_INFO, "main - $section" );
					$tenant=$conf->{$section}->{"tenant"};
					$apitoken = $conf->{$section}->{"token"};
					
					#SHAPES CREATION
					my $shape_plus = $refExcel->add_shape(type => 'mathPlus', id => 1, width => 20, height => 20, fill => '7DCEA0', line => '196F3D', line_weight => 2);
					my $shape_minus = $refExcel->add_shape(type => 'mathMinus', id => 2, width => 20, height => 20, fill => 'CD6155', line => '922B21', line_weight => 2);
					my $shape_equal = $refExcel->add_shape(type => 'mathEqual', id => 3, width => 20, height => 20, fill => '5499C7', line => '1F618D', line_weight => 2);

					#SHEET CREATION
					my $refSheet = $refExcel->add_worksheet($section);

					#SHEET SETTING
            		&paramSheet($refSheet); 

					#HEADER (TENANT STATS)
					my $format = &paramCell($refExcel,"entete");
        			$refSheet->write(0,0, "Month", $format);
        			$refSheet->write(0,1, "DEM - TENANT", $format);
					$refSheet->write(0,2, "Evolution", $format);

					my $title;
					my $compteur=3;
					my $threshold;

					#BROWSING OF THE MANAGEMENT ZONES 
					foreach my $param (sort (keys %{$conf->{$section}})) {
						$threshold=$conf->{$section}->{"limitDEM"};
						if ($param =~ /^ManagementZone/ ){
							$mgtZone{"$param"}= $conf->{$section}->{"limitDEM_$param"};
							#print "\t$param = $conf->{$section}->{$param}\n";
							#push @mgtZone, $conf->{$section}->{$param};
							$title="$conf->{$section}->{$param}";
							&log( L_INFO, "createExcel - $section > Management Zone $title" );

							#REQUESTS SENT TO DYNATRACE API FOR EACH RANGES OF DATE
							#LOOP ON THE MONTHS
							for ( my $i = 0; $i < scalar(@hDateYear); $i++) {
									my %dateYear = %{ $hDateYear[$i] };
									my $refdateYear = \%dateYear;
									getAppList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
									getSyntheticList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
									getHttpMonitorList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
									
									#DISPLAY THE DETAILS (NUMBER OF SESSIONS/ACTIONS BY MONTH) IN A COMMENT
									my $comment="";
									my $totalDEM=&getURL($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"},$title,\$comment);
									$refSheet->write_comment( $i+1, $compteur, $comment, x_scale => 4, y_scale => 4 );

									# HEADER (MANAGEMENT ZONES STATS)
									$format = &paramCell($refExcel,"entete");
									$refSheet->write(0,$compteur, "DEM - ".$title, $format);
									$refSheet->write(0,$compteur+1, "Evolution", $format);
									$format = &paramCell($refExcel,"body");

									# TREND BETWEEN 2 MONTHS
									my $rate="";
									my $cell_format = &paramCell($refExcel,"trend");
									if ($i ne 0){
											# Calculation of the evolution rate
											my $cell_previous_value = xl_rowcol_to_cell( $i, $compteur);
											my $cell_latest_value = xl_rowcol_to_cell( $i+1, $compteur);
											$refSheet->write($i+1,$compteur+1, "=IFERROR( ($cell_latest_value-$cell_previous_value)/$cell_previous_value, IF($cell_latest_value<>0,1,0) )",$cell_format);
											
											# Write a conditional format over a range.
											my $green = $refExcel->add_format( bold => 1 , color => "green", align => 'right' );
											my $red = $refExcel->add_format( bold => 1 , color => "red", align => 'right' );
											my $blue = $refExcel->add_format( bold => 1 , color => "blue", align => 'right' );
											my $cell_start = xl_rowcol_to_cell( 1, $compteur+1);
											my $cell_end = xl_rowcol_to_cell( $i+1, $compteur+1);
											$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '>', value => 0, format => $green });
											$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '<', value => 0, format => $red });
											$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '==', value => 0, format => $blue });
											$refSheet->write($i+1,$compteur, $totalDEM,$format);
									}
									else{
										$cell_format = &paramCell($refExcel,"trend");
										$refSheet->write($i+1,$compteur, $totalDEM,$format);
										$refSheet->write_rich_string($i+1,$compteur+1, "-",$cell_format);
										$refSheet->set_column($compteur+1, $compteur+1, 10);
									}
									$refSheet->set_column(3, $compteur, 22);
							}
							$compteur=$compteur+2;
						}
					}

					# LOOP ON THE MONTHS					
					my $col=1;
					for ( $i = 0; $i < scalar(@hDateYear); $i++) {
						my %dateYear = %{ $hDateYear[$i] };
						my $refdateYear = \%dateYear;
						getAppList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
						getSyntheticList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
						getHttpMonitorList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
						#STYLE OF THE 1ST COLUMN
			        	$format = &paramCell($refExcel,"mois");
						$refSheet->write_date_time($i+1,0, $refdateYear->{"last_month"}."/".$refdateYear->{"last_year"} ,$format);
						#STYLE OF THE OTHER CELLS 
						$format = &paramCell($refExcel,"body");
						#DISPLAY THE DETAILS (NUMBER OF SESSIONS/ACTIONS BY MONTH) IN A COMMENT
						my $comment="";
						my $totalDEM=&getURL($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"},"",\$comment);
						$refSheet->write_comment( $i+1, $col, $comment, x_scale => 4, y_scale => 4);
																							
						#TREND BETWEEN 2 MONTHS
						my $rate="";
						my $cell_format = &paramCell($refExcel,"trend");
						my $cell_start;
						my $cell_end;
						if ($i ne 0){
								# Calculation of the evolution rate
								my $cell_previous_value = xl_rowcol_to_cell( $i, $col);
								my $cell_latest_value = xl_rowcol_to_cell( $i+1, $col);
								$refSheet->write($i+1,$col+1, "=IFERROR( ($cell_latest_value-$cell_previous_value)/$cell_previous_value, IF($cell_latest_value<>0,1,0) )",$cell_format);
								# Write a conditional format over a range.
								my $green = $refExcel->add_format( bold => 1 , color => "green", align => 'right' );
								my $red = $refExcel->add_format( bold => 1 , color => "red", align => 'right' );
								my $blue = $refExcel->add_format( bold => 1 , color => "blue", align => 'right' );
								$cell_start = xl_rowcol_to_cell( 1, $col+1);
								$cell_end = xl_rowcol_to_cell( $i+1, $col+1);
								$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '>', value => 0, format => $green });
								$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '<', value => 0, format => $red });
								$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '==', value => 0, format => $blue });
								$refSheet->write($i+1,$col, $totalDEM,$format);
						}
						else{
							$cell_format = &paramCell($refExcel,"trend");
							$refSheet->write($i+1,$col, $totalDEM,$format);
							$refSheet->write_rich_string($i+1,$col+1, "-",$cell_format);
							$refSheet->set_column($col+1, $col+1, 10);
						}
					}
					$col++;

					#"FOOTER" 
					$format = &paramCell($refExcel,"total");
					$refSheet->write($i+1,0, "TOTAL",$format);
					$format = &paramCell($refExcel,"mois");
					$refSheet->write($i+2,0, "CURRENT POOL LIMIT",$format);
					$refSheet->write($i+3,0, "PERCENT OF USED POOL",$format);
					$format = &paramCell($refExcel,"total");
					$refSheet->write($i+1,1, "=SUM(B2:B".($i+1).")",$format);
					$format = &paramCell($refExcel,"body");
					my $nbMgtZone = keys %mgtZone;
					my $totalLimit=0;
					#IF ONE MANAGEMENT ZONE IS DEFINED
					if ($nbMgtZone ne 0){
						my $k=2;
						foreach my $zone (sort (keys %mgtZone)) {
							$threshold=$mgtZone{$zone};
							$refSheet->write($i+2,$k+1, $threshold,$format);
							
							my $cell_start = xl_rowcol_to_cell( 1, $k+1);
							my $cell_end = xl_rowcol_to_cell( $i, $k+1);
							$format = &paramCell($refExcel,"total");
							$refSheet->write($i+1,$k+1, "=SUM($cell_start:$cell_end)",$format);
							$format = &paramCell($refExcel,"body");
							if ($threshold ne 0){
								$refSheet->write($i+3,$k+1, "=SUM($cell_start:$cell_end)*100/$threshold",$format);
							}
							else{
								$refSheet->write($i+3,$k+1, "N/A",$format);
							}
							$refSheet->write($i+2,$k+1, $threshold,$format);
							$k=$k+2;
						}
					}
					#IF NO MANAGEMENT ZONE IS DEFINED
					$format = &paramCell($refExcel,"body");
					$refSheet->write($i+2,1, $conf->{$section}->{"limitDEM"},$format);
					$refSheet->write($i+3,1, "=SUM(B2:B".($i+1).")*100/".$conf->{$section}->{"limitDEM"},$format);
				}				
			}

    # ADD THE CHART ON THE 1ST EXCEL SHEET
	my $chart = $refExcel->add_chart( type => "line", embedded => 1);
	# ADDING THE LABELS
	$chart->set_title( name => 'DEM Licenses Consumption' );
	$chart->set_x_axis( name => 'Date' );
	$chart->set_y_axis( name => 'Number of DEM' );
	$chart->set_style(2);
	$chart->set_size( width => 1400, height => 400);
	$chart->set_table();

	foreach my $section (keys %{$conf}){
		if ($section =~ /^tenant-/ ){
			$refSheet = $refExcel->get_worksheet_by_name("$section");
            $chart->add_series(
                 categories => '='.$section.'!$A$2:$A$'.($i+1), # Month selection
                 values     => '='.$section.'!$B$2:$B$'.($i+1),  # DEM selection
                 name       => "$section", # Title of the chart
                 marker     => { type => "circle" , size => 7}, # dot marker for each value
                 data_labels => { value => 0,  position => "top" }, # don't display the value in front of the marker
				 gradient => { colors => [ 'red', 'green','blue' ] }
             );
		}
	}
	my $global_worksheet = $refExcel->sheets(0);
	$global_worksheet->insert_chart('A1', $chart);	

}

#---- duplicateExcel -----------------------------------------------------------#
#
# Duplication du precedent fichier d'export dans un nouveau fichier Excel
#
# Entree :
#
# Sortie :
#
# Return :
#

sub duplicateExcel {

my ($refExcel,$refSheet,$refWorkSheet) = @_;

&log( L_INFO, "duplicateExcel - Clone of the worksheet \"".$refSheet->get_name()."\"." );

	my $format = "";	
	foreach my $row (0..$refSheet->{MaxRow}){
        foreach my $col (0..$refSheet->{MaxCol}){
                my $cell = $refSheet -> {Cells} [$row] [$col];
            if ($cell) {
                if ($row == 0){
                      $format = &paramCell($refExcel,"entete");
                      $refWorkSheet->write($row,$col,$refSheet -> {Cells} [$row] [$col]->value,$format);
                }
                elsif ($col == 0){
                      $format = &paramCell($refExcel,"mois");
                      $refWorkSheet->write($row,$col,$refSheet -> {Cells} [$row] [$col]->value,$format);
                }
                elsif ($col != 0 && $col % 2 == 0){
                      $format = &paramCell($refExcel,"trend");
					  $refWorkSheet->write($row,$col, $refSheet -> {Cells} [$row] [$col]{Val}, $format);
			    }
            	else{
                    $format = &paramCell($refExcel,"body");
                    $refWorkSheet->write($row,$col,$refSheet -> {Cells} [$row] [$col]->value,$format);
                }
            }
        }
	}
}

#---- updateExcel --------------------------------------------------------------#
#
# Mise à jour du fichier excel (ajout dernière ligne) si fichier existant
#
# Entree :
#
# Sortie :
#
# Return :
#

sub updateExcel {
my ($refExcel,$refSheet,$refWorkSheet,$refdate) = @_;

# ADD THE LAST LINE
	my $i;
	my $section = $refWorkSheet->get_name();
	if ($section =~ /^tenant-/ ){
	my %mgtZone;
	&log( L_INFO, "updateExcel - $section" );
	$tenant=$conf->{$section}->{"tenant"};
	$apitoken = $conf->{$section}->{"token"};
	
	#SHAPES CREATION
	# not used anymore
	#my $shape_plus = $refExcel->add_shape(type => 'mathPlus', id => 1, width => 20, height => 20, fill => '7DCEA0', line => '196F3D', line_weight => 2);
	#my $shape_minus = $refExcel->add_shape(type => 'mathMinus', id => 2, width => 20, height => 20, fill => 'CD6155', line => '922B21', line_weight => 2);
	#my $shape_equal = $refExcel->add_shape(type => 'mathEqual', id => 3, width => 20, height => 20, fill => '5499C7', line => '1F618D', line_weight => 2);
	#SHEET SELECTION
	my $refSheet = $refExcel->get_worksheet_by_name($section);

	#SHEET SETTING
	&paramSheet($refSheet); 
	
	#HEADER
	my $format = &paramCell($refExcel,"entete");
	$refSheet->write(0,0, "Month", $format);
	$refSheet->write(0,1, "DEM - TENANT", $format);
	$refSheet->write(0,2, "Evolution", $format);
	my $title;
	my $compteur=3;
	my $threshold;
	my $old_value=1;
	foreach my $param (sort (keys %{$conf->{$section}})) {
		$threshold=$conf->{$section}->{"limitDEM"};
		if ($param =~ /^ManagementZone/ ){
		$mgtZone{"$param"}= $conf->{$section}->{"limitDEM_$param"};
		#print "\t$param = $conf->{$section}->{$param}\n";
		#push @mgtZone, $conf->{$section}->{$param};
		$title="$conf->{$section}->{$param}";
		&log( L_INFO, "updateExcel - $section > Management Zone $title" );
			
			#for ( my $i = 0; $i < scalar(@hDateYear); $i++) {
			for ( $i = scalar(@hDateYear)-1; $i < scalar(@hDateYear); $i++) {
					my %dateYear = %{ $hDateYear[$i] };
					my $refdateYear = \%dateYear;
					getAppList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
					getSyntheticList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
					getHttpMonitorList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
					
					#DISPLAY THE DETAILS (NUMBER OF SESSIONS/ACTIONS BY MONTH) IN A COMMENT
					my $comment="";
					my $totalDEM=&getURL($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"},$title,\$comment);
					$refSheet->write_comment( $i+1, $compteur, $comment, x_scale => 4, y_scale => 4 );
					$format = &paramCell($refExcel,"entete");
					$refSheet->write(0,$compteur, "DEM - ".$title, $format);
					$refSheet->write(0,$compteur+1, "Evolution", $format);
					$format = &paramCell($refExcel,"body");
					#TREND BETWEEN 2 MONTHS
					my $rate="";
					my $cell_format = &paramCell($refExcel,"trend");
					if ($i ne 0){
						# Calculation of the evolution rate
						my $cell_previous_value = xl_rowcol_to_cell( $i, $compteur);
						my $cell_latest_value = xl_rowcol_to_cell( $i+1, $compteur);
						$refSheet->write($i+1,$compteur+1, "=IFERROR( ($cell_latest_value-$cell_previous_value)/$cell_previous_value, IF($cell_latest_value<>0,1,0) )",$cell_format);

						# Write a conditional format over a range.
						my $green = $refExcel->add_format( bold => 1 , color => "green", align => 'right' );
						my $red = $refExcel->add_format( bold => 1 , color => "red", align => 'right' );
						my $blue = $refExcel->add_format( bold => 1 , color => "blue", align => 'right' );
						my $cell_start = xl_rowcol_to_cell( 1, $compteur+1);
						my $cell_end = xl_rowcol_to_cell( $i+1, $compteur+1);
						$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '>', value => 0, format => $green });
						$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '<', value => 0, format => $red });
						$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '==', value => 0, format => $blue });
						$refSheet->write($i+1,$compteur, $totalDEM,$format);
					}
					else{
						$cell_format = &paramCell($refExcel,"trend");
						$refSheet->write($i+1,$compteur, $totalDEM,$format);
						$refSheet->write_rich_string($i+1,$compteur+1, "",$cell_format);
						$refSheet->set_column($compteur+1, $compteur+1, 10);
					}
					$refSheet->set_column(3, $compteur, 22);
			}
			$compteur=$compteur+2;
		}
	}
	# LOOP ON THE MONTH -1
	$old_value=1;
	my $col=1;
	
	#for ( $i = 0; $i < scalar(@hDateYear); $i++) {
	for ( $i = scalar(@hDateYear)-1; $i < scalar(@hDateYear); $i++) {
		my %dateYear = %{ $hDateYear[$i] };
		my $refdateYear = \%dateYear;
		getAppList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
		getSyntheticList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
		getHttpMonitorList($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"});
		#STYLE OF THE 1ST COLUMN
    	$format = &paramCell($refExcel,"mois");
		$refSheet->write_date_time($i+1,0, $refdateYear->{"last_month"}."/".$refdateYear->{"last_year"} ,$format);
		#STYLE OF THE OTHER CELLS 
		$format = &paramCell($refExcel,"body");
		#DISPLAY THE DETAILS (NUMBER OF SESSIONS/ACTIONS BY MONTH) IN A COMMENT
		my $comment="";
		my $totalDEM=&getURL($tenant,$apitoken,$refdateYear->{"start"},$refdateYear->{"end"},"",\$comment);
		
		$refSheet->write_comment( $i+1, $col, $comment, x_scale => 4, y_scale => 4);
																			
		#TREND BETWEEN 2 MONTHS
		my $rate="";
		my $cell_format = &paramCell($refExcel,"trend");
		if ($i ne 0){
			# Calculation of the evolution rate
			my $cell_previous_value = xl_rowcol_to_cell( $i, $col);
			my $cell_latest_value = xl_rowcol_to_cell( $i+1, $col);
			$refSheet->write($i+1,$col+1, "=IFERROR( ($cell_latest_value-$cell_previous_value)/$cell_previous_value, IF($cell_latest_value<>0,1,0) )",$cell_format);

			# Write a conditional format over a range.
			my $green = $refExcel->add_format( bold => 1 , color => "green", align => 'right' );
			my $red = $refExcel->add_format( bold => 1 , color => "red", align => 'right' );
			my $blue = $refExcel->add_format( bold => 1 , color => "blue", align => 'right' );
			my $cell_start = xl_rowcol_to_cell( 1, $col+1);
			my $cell_end = xl_rowcol_to_cell( $i+1, $col+1);
			$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '>', value => 0, format => $green });
			$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '<', value => 0, format => $red });
			$refSheet->conditional_formatting( "$cell_start:$cell_end",{ type => 'cell', criteria => '==', value => 0, format => $blue });
			$refSheet->write($i+1,$col, $totalDEM,$format);
		}
		else{
			$cell_format = &paramCell($refExcel,"trend");
			$refSheet->write($i+1,1, $totalDEM,$format);
			$refSheet->write_rich_string($i+1,$col+1, "-",$cell_format);
		}		
	}
	$col++;
	#"FOOTER" 
	$format = &paramCell($refExcel,"total");
	$refSheet->write($i+1,0, "TOTAL",$format);
	$format = &paramCell($refExcel,"mois");
	$refSheet->write($i+2,0, "CURRENT POOL LIMIT",$format);
	$refSheet->write($i+3,0, "PERCENT OF USED POOL",$format);
	$format = &paramCell($refExcel,"total");
	$refSheet->write($i+1,1, "=SUM(B2:B".($i+1).")",$format);
	$format = &paramCell($refExcel,"body");
	my $nbMgtZone = keys %mgtZone;
	my $totalLimit=0;
	#IF ONE MANAGEMENT ZONE IS DEFINED
	if ($nbMgtZone ne 0){
		my $k=2;
		foreach my $zone (sort (keys %mgtZone)) {
			$threshold=$mgtZone{$zone};
			$refSheet->write($i+2,$k+1, $threshold,$format);
			
			my $cell_start = xl_rowcol_to_cell( 1, $k+1);
			my $cell_end = xl_rowcol_to_cell( $i, $k+1);
			$format = &paramCell($refExcel,"total");
			$refSheet->write($i+1,$k+1, "=SUM($cell_start:$cell_end)",$format);
			$format = &paramCell($refExcel,"body");
			if ($threshold ne 0){
				$refSheet->write($i+3,$k+1, "=SUM($cell_start:$cell_end)*100/$threshold",$format);
			}
			else{
				$refSheet->write($i+3,$k+1, "N/A",$format);
			}
			#$totalLimit=$totalLimit+$threshold;
			$refSheet->write($i+2,$k+1, $threshold,$format);
			$k=$k+2;
		}
	}
	$format = &paramCell($refExcel,"body");
	$refSheet->write($i+2,1, $conf->{$section}->{"limitDEM"},$format);
	$refSheet->write($i+3,1, "=SUM(B2:B".($i+1).")*100/".$conf->{$section}->{"limitDEM"},$format);
	}
}



#---- makeChart --------------------------------------------------------------#
#
# Build the chart on the first worksheet
#
# Entree :
#
# Sortie :
#
# Return :
#

sub makeChart {
	my ($refExcel,$refSheet,$refWorkSheet,$refdate) = @_;
	&log( L_INFO, "makeChart - Build of the chart on the first sheet" );
    #my $sheet = @{$refExcel -> {Worksheet}}[1];
	my $last_row=($refSheet->{MaxRow})-3;

	my $chart = $refExcel->add_chart( type => "line", embedded => 1);
	# Ajout des labels
	$chart->set_title( name => 'DEM Licenses Consumption' );
	$chart->set_x_axis( name => 'Date' );
	$chart->set_y_axis( name => 'Number of DEM' );
	$chart->set_style(2);
	$chart->set_size( width => 1400, height => 400);
	$chart->set_table();

	foreach my $section (keys %{$conf}){
		if ($section =~ /^tenant-/ ){
			$refWorkSheet = $refExcel->get_worksheet_by_name("$section");

            $chart->add_series(
                 categories => '='.$section.'!$A$2:$A$'.($last_row+2), # Month selection
                 values     => '='.$section.'!$B$2:$B$'.($last_row+2),  # DEM selection
                 name       => "$section", # Title of the chart
                 marker     => { type => "circle" , size => 7}, # dot marker for each value
                 data_labels => { value => 0,  position => "top" }, # don't display the value in front of the marker
				 gradient => { colors => [ 'red', 'green','blue' ] }
             );
		}
	}
	$refWorkSheet = $refExcel->sheets(0);
	$refWorkSheet->insert_chart('A1', $chart);
}