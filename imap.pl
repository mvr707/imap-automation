#!/usr/bin/env perl

use strict; 
use warnings;

use CGI;
use Mail::IMAPClient;
use DateTime::Format::Mail;
use Date::Manip;
use Data::Table;
use Email::MIME;
use Data::Dumper;
use Email::MIME::Attachment::Stripper;
use Regexp::Common;
use Term::ReadLine;

my $q_current = {};

if (-e "$0.conf") {
	open(my $fh, "$0.conf");
	$q_current = { %$q_current, %{ CGI->new($fh)->Vars } };
	close($fh);
}

if (-e "$ENV{HOME}/.imaprc") {
	open(my $fh, "$ENV{HOME}/.imaprc");
	$q_current = { %$q_current, %{ CGI->new($fh)->Vars } };
	close($fh);
}

$q_current = { %$q_current, %{ CGI->new()->Vars } };

my $client;

use Term::ReadKey;

sub collect_password 
{
	$q_current->{password} = eval {
		print "Enter Password : \n";
		ReadMode 4; # Turn off controls keys
		my $p = ReadLine(0);
		ReadMode 0; # Reset tty mode before exiting
		chomp($p);
		return $p;
	}
}

if (!$q_current->{password} or $q_current->{password} eq '-') {
	collect_password();
}


sub connect_with_server()
{
	# Connect to IMAP server
	$client = Mail::IMAPClient->new(
		Server   => $q_current->{server},
		User     => $q_current->{user},
		Password => ($q_current->{password} =~ /^[01]*$/ ? pack('B*', $q_current->{password}) : $q_current->{password}),
		Port     => 993,
		Ssl      =>  1,
		IgnoreSizeErrors => 1,	### Certain (caching) servers, like Exchange 2007, often report the wrong message size. 
					### Instead of chopping the message into a size that it fits the specified size, 
					### the reported size will be simply ignored when this parameter is set to 1.
	) or die "ERROR: Cannot connect through IMAPClient: $!\n";

}

connect_with_server();




my $results; ### Current search results;

### Folder and Search be specified for a sumamry
sub print_summary()
{
	my $folder = $q_current->{folder} or die "no folder specified\n";
	my $search = $q_current->{search};

	$client->select($folder) or die "ERROR: Can not select folder=$folder " . $client->LastError . "\n";
	$results = $client->search($search) or do { print "can not search '$search' " . $client->LastError . "\n"; return; };
	my $hr = $client->parse_headers($results, 'From', 'To', 'Subject', 'Date');
	#print Dumper($hr);

	my $data = [];

	for my $i (sort keys %$hr) {
		my $ddd = $hr->{$i}{Date}[0];
		$ddd =~ s/\(.*?\)//g;
		my $ddd_std = eval { DateTime::Format::Mail->parse_datetime($ddd) } || $ddd;
		push(@$data, [$folder, $i, $ddd_std, $hr->{$i}{From}[0], $hr->{$i}{To}[0], $hr->{$i}{Subject}[0]]);
	}
	my $dt = Data::Table->new($data, [qw/Folder ID Date From To Subject/], 0);
	$dt->sort('Date', 1, 0);
	print $dt->tsv;
}

if ($q_current->{folder} and $q_current->{search}) {
	print_summary();
}

my $action = $q_current->{'action'} || '';

my $batch_size = $q_current->{'batch_size'} || 1000;

if ($action =~ /move_to_mail_folder:(\w+)/) {
	my $move_to_mail_folder = $1;

	my $count = @$results;
	print "\nMoving $count messages to folder=$move_to_mail_folder\n";

	while (0 != @$results) {
		my $batch = [ splice(@$results, 0, $batch_size) ];
		my $batch_count = @$batch;
		$client->move($move_to_mail_folder, $batch) or die "ERROR: Can not move_to_mail_folder:$move_to_mail_folder " . $client->LastError . "\n"; 
		print "\t $batch_count messages moved to folder=$move_to_mail_folder\n";
	}
	print "$count messages moved to folder=$move_to_mail_folder\n";

#	for my $i (@$results) {
#		$client->move($move_to_mail_folder, $i) or die "ERROR: Can not move_to_mail_folder:$move_to_mail_folder " . $client->LastError . "\n"; 
#		print "\t $i moved to folder=$move_to_mail_folder\n";
#	}

	$client->expunge();
	exit(0);
}

if ($action =~ /copy_to_local_folder:(\S+)/) {
	my $copy_to_local_folder = $1;
	my $folder = $q_current->{folder};
	for my $i (@$results) {
		$client->message_to_file("$copy_to_local_folder/$folder.$i.mime", $i) or die "ERROR: Can not copy_to_local_folder:$copy_to_local_folder " . $client->LastError . "\n";
		print "\t $i copied to local $copy_to_local_folder/$folder.$i.mime\n";
	}
	exit(0);
}

if ($action =~ /copy_to_local_file:(\S+)/) {
	my $copy_to_local_file = $1;
	$client->message_to_file($copy_to_local_file, @$results) or die "ERROR: Can not copy_to_local_file:$copy_to_local_file " . $client->LastError . "\n";
	print "\t Results copied to local file $copy_to_local_file\n";
	exit(0);
}


if ($action ne 'shell') {
	exit(0);
}


my $term = Term::ReadLine->new('IMAP Shell');
my $prompt = derive_prompt();

my $line;

my $body_cached = {};

while ( defined ($line = $term->readline($prompt)) ) {
	my $q1_hash_ref = CGI->new(join(';', map { s/'//g; s/"//g; $_ } ($line =~ /\S*$RE{quoted}\S*|\S+/g)))->Vars;

	$q_current =  { %$q_current, %$q1_hash_ref };


	if ($q1_hash_ref->{folder} or $q1_hash_ref->{search}) {
		print_summary();
		delete $q_current->{read};
	}

	if ($q_current->{folder} and !$q_current->{search}) {
		print "'search' not specified. You may specify search=ALL..\n";
		if ($client->selectable($q_current->{folder})) {
			my $message_count = $client->message_count($q_current->{folder}) || 0;
			print "\t* $q_current->{folder} ($message_count)\n";
		} else {
			print "WARNING: folder=$q_current->{folder} is not selectable\n";
		}
	}

	if (!$q_current->{folder}) {
		print "'folder' not specified. Please specify one...\n";
		print "folders:\n";
		for my $i ($client->folders()) {
			next if (!$client->selectable($i));
			my $message_count = $client->message_count($i) || 0;
			print "\t* $i (", $message_count, ")\n";
		}
	}

	if ($q_current->{quit}) {
			exit(0);
	}

	if ($q_current->{read}) {
		$body_cached = {};
		my $msg_id = $q_current->{read};
		my $msg = [ $msg_id ];
		my $hr = $client->parse_headers($msg, 'From', 'To', 'Subject', 'Date');
		my $ddd = $hr->{$msg_id}{Date}[0];
		$ddd =~ s/\(.*?\)//g;
		my $ddd_std = eval { DateTime::Format::Mail->parse_datetime($ddd) } || $ddd;

		my $mime = $client->message_string($msg_id) || print "\n\nERROR: Can not read $msg_id\n\n" && next;
		my $body_string = $client->body_string($msg_id) || print "\n\nERROR: Can not read $msg_id\n\n" && next;

		my $parsed = Email::MIME->new($mime);
		my $message_structure = $parsed->debug_structure;
		my $decoded = $parsed->body;
		my @parts = $parsed->parts;

		my @ct = map { $_->content_type } @parts;

		my $stripper = Email::MIME::Attachment::Stripper->new($mime);
		my $message = $stripper->message;
		my $number_of_attachments = scalar($stripper->attachments);

		print <<EOF;
Date: $ddd_std
From: $hr->{$msg_id}{From}[0]
To: $hr->{$msg_id}{To}[0]
Subject: $hr->{$msg_id}{Subject}[0]

NumberOfAttachments: $number_of_attachments
MessageStructure: 
$message_structure

EOF
		traverse_mime($parsed, "1", 0);

		for my $i (sort keys %$body_cached) {
		print <<EOF;

[$i]
$body_cached->{$i}
[/$i]
EOF
		}

		if ($q_current->{detail}) {
			print $mime;
		}
		#delete $q_current->{read};
	}

	$prompt = derive_prompt();
}

sub derive_prompt
{
        #return CGI->new($current)->query_string;

        my $tmp = '';   
        $tmp = "[$q_current->{folder}" if ($q_current->{folder});
        if ($tmp) {
                if ($q_current->{search}) {
                        $tmp .= "|$q_current->{search}";
			if ($q_current->{read}) {
				$tmp .= "|$q_current->{read}";
			}
                }
                $tmp .= "]";
        }
        $tmp = ($tmp ? "$tmp >> " : ">> ");

        return $tmp;
}

sub traverse_mime
{
	my $in = shift;
	my $label = shift;
	my $depth = shift;

	my $ct = $in->content_type;

	my @parts = $in->parts(); 
	my $num_parts = @parts;
	my @ct_parts = map { $_->content_type } @parts;
	my $ct_parts = join(',', @ct_parts);

	my @sub_parts = $in->subparts();
	my $num_subparts = @sub_parts;
	my @ct_subparts = map { $_->content_type } @sub_parts;
	my $ct_subparts = join(',', @ct_subparts);

	print "\t" x $depth, "$label) $ct\n"; 

	if ($ct =~ /text\/plain/) {
		$body_cached->{$label} = $in->body;
	}

	my $tmp = 1;
	for my $p (@sub_parts) {
		traverse_mime($p, "$label.$tmp", $depth+1);
		$tmp++;
	}
}
