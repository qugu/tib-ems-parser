#!/usr/bin/perl
use strict;
use warnings;
use Data::Dumper;
use JSON::Parse ':all';

use Excel::Writer::XLSX;
use List::Compare;

# Check for command line arguments before running
my $args_num = $#ARGV + 1;
if ($args_num != 2 ) {
	print "\nUsage: $0 file.json file.xlsx \n";
	exit;
		}

my ($input_file, $result_file) = @ARGV;

# JSON to perl conversion
my $json = json_file_to_perl ($input_file);
# JSON to perl conversion - "new" file;
my $json2 = json_file_to_perl ("json_users_short_2.json");



print Dumper @{$json->{users}};

# my $lc = List::Compare->new(\@{$json->{users}},\@{$json2->{users}});
# my @intersection = $lc->get_intersection;

# print Dumper \@intersection;

#
# Iterate through users
#

# # Row reset;
# my $row = 1;

# # Old code for now
# foreach my $entry ( sort { $a->{name} cmp $b->{name} } @{ $json->{users}}) {
   # my $field1 = $entry->{'name'};
   # my $field2 = $entry->{'description'};
      
   	# # Set xls table column to zero each iteration of the loop
	# my $col = 0;
		# # Loop to write values in xls file 
	# $users_ws->write ( $row, $col, $field1, $format2 );
		# $col++;
		# $users_ws->write ( $row, $col, $field2, $format2 );
			# $row++;
   
   # }
   
# # Reset row variable for new iterations
# $row = 1;

# #
# # Iterate through topics
# #
# foreach my $entry (sort { $a->{name} cmp $b->{name} } @{ $json->{topics}}) {
   # my $field1 = $entry->{'name'};
   # my $field2 = $entry->{'store'};
   # my $field3 = $entry->{'trace'};
   # my $field4 = $entry->{'maxbytes'};
   # my $field5 = $entry->{'secure'};
      
   	# # Set xls table column to zero each iteration of the loop
	# my $col = 0;
		# # Loop to write values in xls file 
	# $topics_ws->write ( $row, $col, $field1, $format2 ); $col++;
	# $topics_ws->write ( $row, $col, $field2, $format2 ); $col++;
	# $topics_ws->write ( $row, $col, $field3, $format2 ); $col++;
	# $topics_ws->write ( $row, $col, $field4, $format2 ); $col++;
	# $topics_ws->write ( $row, $col, $field5, $format2 ); $col++;
			# $row++;
   
   # }   

# # Reset row var for xls
# $row = 1;

# # Iterate through queues
# foreach my $entry (sort { $a->{name} cmp $b->{name} } @{ $json->{queues}}) {
   # my $field1 = $entry->{'name'};
   # my $field2 = $entry->{'store'};
   # my $field3 = $entry->{'trace'};
   # my $field4 = $entry->{'maxbytes'};
   # my $field5 = $entry->{'secure'};
      
   	# # Set xls table column to zero each iteration of the loop
	# my $col = 0;
		# # Loop to write values in xls file 
	# $queues_ws->write ( $row, $col, $field1, $format2 ); $col++;
	# $queues_ws->write ( $row, $col, $field2, $format2 ); $col++;
	# $queues_ws->write ( $row, $col, $field3, $format2 ); $col++;
	# $queues_ws->write ( $row, $col, $field4, $format2 ); $col++;
	# $queues_ws->write ( $row, $col, $field5, $format2 ); $col++;
			# $row++;
   
   # }   

# # Will be writing to excel here in the end;
# sub_excel_create {
# ## Create a new Excel workbook
# my $workbook = Excel::Writer::XLSX->new($result_file);

# # Add formatting for the worksheet titles
# my $format1 = $workbook->add_format();
# $format1->set_format_properties ( bold => '1', color => 'black', align => 'left' ) ;
# my $format2 = $workbook->add_format();
# $format2->set_format_properties ( font => 'Calibri', size => '10' );

# # Add a worksheet and add basic information to it;
# my $users_ws = $workbook->add_worksheet('Users');
# my $topics_ws = $workbook->add_worksheet('Topics');
# my $queues_ws = $workbook->add_worksheet('Queues');

# $users_ws->write ( 0, 0, 'Username', $format1);
# $users_ws->write ( 0, 1, 'Description from EMS', $format1 );

# $topics_ws->write ( 0, 0, 'Name', $format1 ); 
# $topics_ws->write ( 0, 1, 'Store', $format1 );
# $topics_ws->write ( 0, 2, 'Trace', $format1 );
# $topics_ws->write ( 0, 3, 'Maxbytes', $format1 );
# $topics_ws->write ( 0, 4, 'Secure', $format1 );

# $queues_ws->write ( 0, 0, 'Name', $format1 ); 
# $queues_ws->write ( 0, 1, 'Store', $format1 );
# $queues_ws->write ( 0, 2, 'Trace', $format1 );
# $queues_ws->write ( 0, 3, 'Maxbytes', $format1 );
# $queues_ws->write ( 0, 4, 'Secure', $format1 );

 # # Set column' widths
# $users_ws->set_column( 0, 0, 40 );
# $users_ws->set_column( 'B:C', 35 );

# $topics_ws->set_column( 0, 0, 40 );
# $topics_ws->set_column( 'B:F', 20 );
# $queues_ws->set_column( 0, 0, 40 );
# $queues_ws->set_column( 'B:F', 20 );

# } 

   
# $workbook->close();
