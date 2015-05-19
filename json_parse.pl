#!/usr/bin/perl
use strict;
use warnings;
use Data::Dumper;
use JSON::Parse ':all';
use Excel::Writer::XLSX;

# JSON to perl conversion
my $json = json_file_to_perl ('tibemsd.json');

# my $file_name = $json->{'server'};
# print Dumper $file_name;

## Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( "excel.xlsx" );

# Add formatting for the worksheet titles
my $format1 = $workbook->add_format();
$format1->set_format_properties ( bold => '1', color => 'black', align => 'left' ) ;
my $format2 = $workbook->add_format();
$format2->set_format_properties ( font => 'Calibri', size => '10' );

# Add a worksheet and add basic information to it;
my $users_ws = $workbook->add_worksheet('Users');
my $topics_ws = $workbook->add_worksheet('Topics');
my $queues_ws = $workbook->add_worksheet('Queues');

$users_ws->write ( 0, 0, 'Username', $format1);
$users_ws->write ( 0, 1, 'Description from EMS', $format1 );
$topics_ws->write ( 0, 0, 'Name', $format1 ); 
$topics_ws->write ( 0, 1, 'Store', $format1 );
$topics_ws->write ( 0, 2, 'Trace', $format1 );
$topics_ws->write ( 0, 3, 'Maxbytes', $format1 );
$topics_ws->write ( 0, 4, 'Secure', $format1 );

$queues_ws->write ( 0, 0, 'Name', $format1 ); 
$queues_ws->write ( 0, 1, 'Store', $format1 );
$queues_ws->write ( 0, 2, 'Trace', $format1 );
$queues_ws->write ( 0, 3, 'Maxbytes', $format1 );
$queues_ws->write ( 0, 4, 'Secure', $format1 );

 # Set column' widths
$users_ws->set_column( 0, 0, 40 );
$users_ws->set_column( 'B:C', 35 );

$topics_ws->set_column( 0, 0, 40 );
$topics_ws->set_column( 'B:F', 20 );
$queues_ws->set_column( 0, 0, 40 );
$queues_ws->set_column( 'B:F', 20 );


my $row = 1;

# Iterate through users
foreach my $entry ( sort { $a->{name} cmp $b->{name} } @{ $json->{users}}) {
   my $field1 = $entry->{'name'};
   my $field2 = $entry->{'description'};
      
   	# Set xls table column to zero each iteration of the loop
	my $col = 0;
		# Loop to write values in xls file 
	$users_ws->write ( $row, $col, $field1, $format2 );
		$col++;
		$users_ws->write ( $row, $col, $field2, $format2 );
			$row++;
   
   }
   
# Reset row var for xls
$row = 1;

# Iterate through topics
foreach my $entry (sort { $a->{name} cmp $b->{name} } @{ $json->{topics}}) {
   my $field1 = $entry->{'name'};
   my $field2 = $entry->{'store'};
   my $field3 = $entry->{'trace'};
   my $field4 = $entry->{'maxbytes'};
   my $field5 = $entry->{'secure'};
      
   	# Set xls table column to zero each iteration of the loop
	my $col = 0;
		# Loop to write values in xls file 
	$topics_ws->write ( $row, $col, $field1, $format2 ); $col++;
	$topics_ws->write ( $row, $col, $field2, $format2 ); $col++;
	$topics_ws->write ( $row, $col, $field3, $format2 ); $col++;
	$topics_ws->write ( $row, $col, $field4, $format2 ); $col++;
	$topics_ws->write ( $row, $col, $field5, $format2 ); $col++;
			$row++;
   
   }   

# Reset row var for xls
$row = 1;

# Iterate through topics
foreach my $entry (sort { $a->{name} cmp $b->{name} } @{ $json->{queues}}) {
   my $field1 = $entry->{'name'};
   my $field2 = $entry->{'store'};
   my $field3 = $entry->{'trace'};
   my $field4 = $entry->{'maxbytes'};
   my $field5 = $entry->{'secure'};
      
   	# Set xls table column to zero each iteration of the loop
	my $col = 0;
		# Loop to write values in xls file 
	$queues_ws->write ( $row, $col, $field1, $format2 ); $col++;
	$queues_ws->write ( $row, $col, $field2, $format2 ); $col++;
	$queues_ws->write ( $row, $col, $field3, $format2 ); $col++;
	$queues_ws->write ( $row, $col, $field4, $format2 ); $col++;
	$queues_ws->write ( $row, $col, $field5, $format2 ); $col++;
			$row++;
   
   }   
   
$workbook->close();

	# foreach my $key (keys %$p) {
	
	
		# my $array_num = 0;
		# print  "$p->{'users'}[$array_num]{name}\n";
		# $array_num++;
		# print "$key: \n";
	# }
	