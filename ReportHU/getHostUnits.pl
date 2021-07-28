#!/bin/perl -w
#--------------------------------------------------------------------------------------------------------------------------------------------------------------#
# AUTHOR        : Ackermann Matthieu
# DATE          : 2021-07-23
# DESCRIPTION   : Export in an Excel file of the Dynatrace Host Units consumption (Oneagent licenses) in SaaS mode 
# UPDATE        : 
#--------------------------------------------------------------------------------------------------------------------------------------------------------------#
use strict;
use LWP::UserAgent;
#use Data::Dumper;
use JSON;
use JSON::XS qw(encode_json);
use Config::Tiny;
use DateTime;

#EXCEL MODULES
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;

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
$conf = Config::Tiny->read( "getHostUnits.ini" );

# INITIALIZATION LOG FILE
my $LOG_FILE = $conf->{"global"}->{"logfile"};
&log( L_INFO, "main - Start $0");

#VARIABLES
my $exportFile = $conf->{"global"}->{"exportfileHU"};
my $tenant;
my $apitoken;
my $url;
my %hostList;
my %mgtZones;
my $currentDate=DateTime->now->ymd('');

getOption();

#DECLARATION OF LWP USER AGENT
my $ua = LWP::UserAgent->new;
$ua->timeout(100);

my $excel_w     = Excel::Writer::XLSX->new(substr($exportFile,0,length($exportFile)-5)."_".$currentDate.".xlsx");
&createExcel($excel_w);
$excel_w->close();
#print Dumper(%mgtZones);
&log( L_INFO, "main - End of script " );

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
	#chop (my $Date = `/bin/date "+[%x %H:%M:%S]"`);
	my $dt = DateTime->now;
	my $ymd = $dt->ymd;
	my $hms = $dt->hms;
	my $Date = "[".$ymd." ".$hms."]";

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

	EXEMPLE:
	perl $0 \n
	\n";
	exit 0;
}
# GETTING OPTIONS
sub getOption {
	use Getopt::Long;
	my $help;
	GetOptions ("help" => \$help)
		or die("Error in command line arguments\n");

	if (defined $help) { 
		&usage(); 
	}
}

# GETTING HOST LIST
sub getHostList{
	my $tenant = $_[0];
	my $token = $_[1];
	my $url="https://$tenant.live.dynatrace.com/api/v1/entity/infrastructure/hosts?relativeTime=day";
	&log( L_INFO, "getHostList - $url" );
	my $response = $ua->get($url, "Content-Type" => "application/json", "Authorization" => "Api-Token $token");

	if ( $response->is_success ) {
		&log( L_INFO, "getHostList - Response : ".$response->status_line );
		my $DT_JSON=$response->decoded_content;
		my $decoded_json = decode_json($DT_JSON);
		
		foreach my $item (sort @$decoded_json) {
			$hostList{$tenant}{$item->{'displayName'}}{'lastSeenTimestamp'} = $item->{'lastSeenTimestamp'};
			$hostList{$tenant}{$item->{'displayName'}}{'consumedHostUnits'} = $item->{'consumedHostUnits'};
			$hostList{$tenant}{$item->{'displayName'}}{'osType'} = $item->{'osType'};
			$hostList{$tenant}{$item->{'displayName'}}{'agentVersion'} = $item->{'agentVersion'}->{'major'}.".".$item->{'agentVersion'}->{'minor'}.".".$item->{'agentVersion'}->{'revision'};
			$hostList{$tenant}{$item->{'displayName'}}{'osVersion'} = $item->{'osVersion'};
			$hostList{$tenant}{$item->{'displayName'}}{'monitoringMode'} = $item->{'monitoringMode'};

			# TAGS INSIDE AN ARRAY OF HASH 
			my $tags = $item->{'tags'};
			my $nbTags = keys @$tags;
			my $tagList="";
			for (my $i=0;$i<$nbTags-1;$i++){
				$tagList .= $item->{'tags'}->[$i]->{'key'}."/";
			}
			if ($tagList ne ""){
				chop($tagList);
				$hostList{$tenant}{$item->{'displayName'}}{'tags'} = $tagList;
			}
			else{
				$hostList{$tenant}{$item->{'displayName'}}{'tags'} = "N/A";
			}

			# MANAGEMENT ZONES INSIDE AN ARRAY OF HASH
			my $mgtZone = $item->{'managementZones'};
			my $nbMgtZones = keys @$mgtZone;
			my $mgtZoneList="";
			for (my $i=0;$i<$nbMgtZones;$i++){
				$mgtZoneList .= $item->{'managementZones'}->[$i]->{'name'}."/";
				$mgtZones{$tenant}{$item->{'managementZones'}->[$i]->{'name'}}+=$item->{'consumedHostUnits'};
			}
			if ($mgtZoneList ne ""){
				chop($mgtZoneList);
				$hostList{$tenant}{$item->{'displayName'}}{'managementZones'} = $mgtZoneList;
			}
			else{
				$hostList{$tenant}{$item->{'displayName'}}{'managementZones'} = "N/A";
				$mgtZones{$tenant}{"N/A"}+=$item->{'consumedHostUnits'};
			}
		}
	}
	else {
		&log( L_ERROR, "getHostList - $url");
		&log( L_ERROR, "getHostList - Error: ".$response->status_line );
		&log( L_ERROR, "getHostList - Error: ".$response->content );
		return ERROR;
	}
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
	$refSheet->set_column(1,1,13);
	# Zoom to 85% (best view)
	$refSheet->set_zoom(85);
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

	elsif ( $type eq 'entete2' ){
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
                        align => 'vcenter',
                        );
        }

#	&log( L_INFO, "paramCell - \$format = |$format|" );
	return $format;
}

#---- createExcel -----------------------------------------------------------#
#
# Creation of the Excel file 
#
# Entree :
#
# Sortie :
#
# Return :
#

sub createExcel {
    my $refExcel        = $_[0];
            #FILE CREATION
			&log( L_INFO, "createExcel - Excel file creation" );
			my $i;
			foreach my $section (sort (keys %{$conf})){
				if ($section =~ /^tenant-/ ){
					my %mgtZone;
					$tenant=$conf->{$section}->{"tenant"};
					$apitoken = $conf->{$section}->{"token"};
					getHostList($tenant,$apitoken);

					#SHEET CREATION
					my $refSheet = $refExcel->add_worksheet($section);
					#SHEET SETTING
            				&paramSheet($refSheet);
					#HOST UNIT CONSUMPTION PER SERVER
					#HEADER
					my $format = &paramCell($refExcel,"entete");
	        			$refSheet->write(0,0, "Servername", $format);
        				$refSheet->write(0,1, "lastSeenTimestamp", $format);
					$refSheet->write(0,2, "osType", $format);
					$refSheet->write(0,3, "osVersion", $format);
					$refSheet->write(0,4, "agentVersion", $format);
					$refSheet->write(0,5, "monitoringMode", $format);
					$refSheet->write(0,6, "consumedHostUnits", $format);
					$refSheet->write(0,7, "tags", $format);
					$refSheet->write(0,8, "managementZones", $format);	

					my $i=0;
					my %monitoringMode;
					$monitoringMode{"INFRASTRUCTURE"}{"count"}=0;
					$monitoringMode{"FULLSTACK"}{"count"}=0;

					&log( L_INFO, "createExcel - Data writing" );
					foreach my $servername (sort keys %{$hostList{$tenant}}){
						# Conversion from timestamp ms to datetime
						use POSIX qw(strftime);
						my $lastSeenTimestamp = strftime "%F %H:%M:%S ", localtime($hostList{$tenant}{$servername}->{"lastSeenTimestamp"}/1000);
						
						my $osType=$hostList{$tenant}{$servername}->{"osType"};
						my $osVersion=$hostList{$tenant}{$servername}->{"osVersion"};
						my $agentVersion=$hostList{$tenant}{$servername}->{"agentVersion"};
						my $monitoringMode=$hostList{$tenant}{$servername}->{"monitoringMode"};
						my $consumedHostUnits=$hostList{$tenant}{$servername}->{"consumedHostUnits"};

						if ($monitoringMode eq "INFRASTRUCTURE"){
							$monitoringMode{"INFRASTRUCTURE"}{"count"}+=1;
							$monitoringMode{"INFRASTRUCTURE"}{"totalhu"}+=$consumedHostUnits;
						}
						else{
							$monitoringMode{"FULLSTACK"}{"count"}+=1;
							$monitoringMode{"FULLSTACK"}{"totalhu"}+=$consumedHostUnits;
						}

						my $tags=$hostList{$tenant}{$servername}->{"tags"};
						my $managementZones=$hostList{$tenant}{$servername}->{"managementZones"};

						#STYLE OF THE CELLS 
						$format = &paramCell($refExcel,"body");
						$refSheet->set_column(0,0,47);
						$refSheet->write($i+1,0, $servername,$format);
						$refSheet->set_column(1,1,19);
						$refSheet->write($i+1,1, $lastSeenTimestamp,$format);
						$refSheet->set_column(2,2,10);
						$refSheet->write($i+1,2, $osType,$format);
						$refSheet->set_column(3,3,70);
						$refSheet->write($i+1,3, $osVersion,$format);
						$refSheet->set_column(4,4,13);
						$refSheet->write($i+1,4, $agentVersion,$format);
						$refSheet->set_column(5,5,16);
						$refSheet->write($i+1,5, $monitoringMode,$format);
						$refSheet->set_column(6,6,18);
						$refSheet->write($i+1,6, $consumedHostUnits,$format);
						$refSheet->set_column(7,7,30);
						$refSheet->write($i+1,7, $tags,$format);
						$refSheet->set_column(8,8,30);
						$refSheet->write($i+1,8, $managementZones,$format);
						$i++;
					}

					#HOST UNIT CONSUMPTION PER MANAGEMENT ZONE 
					$format = &paramCell($refExcel,"entete");
					$refSheet->write($i+2,0, "Management Zone", $format);
					$refSheet->write($i+2,1, "Consumed Host Units", $format);

					$format = &paramCell($refExcel,"body");
					my $j=$i;
					foreach my $mgtZone (sort keys %{$mgtZones{$tenant}}){
						$refSheet->write($j+3,0, $mgtZone,$format);
						$refSheet->write($j+3,1, $mgtZones{$tenant}{$mgtZone},$format);
						$j++;
					}
					$format = &paramCell($refExcel,"entete2");
					$refSheet->write($j+4,0, "TOTAL MONITORED SERVERS",$format);
					$refSheet->write($j+5,0, "TOTAL CONSUMED HOST UNITS",$format);

					$format = &paramCell($refExcel,"body");
					$refSheet->write($j+4,1, "=COUNT(G2:G".($i+1).")",$format);
					$refSheet->write($j+5,1, "=SUM(G2:G".($i+1).")",$format);

					#my $cell_totalhu = $refSheet->get_cell($j+5,1);
					# FILTER CREATION ON THE 1ST ROW
					$refSheet->autofilter( 0, 0, 0, 8 );	

					#HOST UNIT CONSUMPTION PER MONITORING MODE 
					$format = &paramCell($refExcel,"entete");
					$refSheet->write($i+2,3, "Monitoring Mode", $format);
					$refSheet->write($i+2,4, "Total", $format);
					$refSheet->write($i+2,5, "Host Units", $format);
					
					$format = &paramCell($refExcel,"body");
					foreach my $mode (sort keys %monitoringMode){
						$refSheet->write($i+3,3, $mode,$format);
						$refSheet->write($i+3,4, $monitoringMode{$mode}{"count"},$format);
						$refSheet->write($i+3,5, $monitoringMode{$mode}{"totalhu"},$format);
						$i++;
					}
					
					# ADDING CHART
					&log( L_INFO, "createExcel - Chart insertion" );
					my $chart = $refExcel->add_chart( type => "pie", embedded => 1);
					my $refWorkSheet = $refExcel->get_worksheet_by_name("$section");
					$chart->add_series(
						name       => 'Host Units Consumption depending on Monitoring Mode',
						categories => [ "$section", $i+1, $i+2, 3, 3 ],
						values     => [ "$section", $i+1, $i+2, 5, 5 ],
						data_labels => { value => 1, leader_lines => 1},
					);
					$refWorkSheet->insert_chart($i+4,3, $chart);
				}				
			}
			&log( L_INFO, "createExcel : Excel file \"$exportFile\" generated in the current directory" );
}

# GET Host Unit Consumption per server (command-line)
#curl -s -X GET -H "header=present; charset=utf-8" -H "Authorization:Api-Token ****************" "https://********.live.dynatrace.com/api/v1/entity/infrastructure/hosts?relativeTime=3days" | jq -r '.[] | [.displayName, .monitoringMode, .consumedHostUnits, (.managementZones[]? | .name)]|@csv'

