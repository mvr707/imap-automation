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
use File::Path qw(make_path);

### Initialize the default config
my $config = {};

### Override the initial config with any deployment site specific info
if (-e "$0.conf") {
        open(my $fh, "$0.conf");
        $config = { %$config, %{ CGI->new($fh)->Vars } };
        close($fh);
}

### Override the deployment site specific config with any user specific info
if (-e "$ENV{HOME}/.imaprc") {
        open(my $fh, "$ENV{HOME}/.imaprc");
        $config = { %$config, %{ CGI->new($fh)->Vars } };
        close($fh);
}

### Now, override with any commandline supplied config
$config = { %$config, %{ CGI->new()->Vars } };

### Config: Override with values with specific config ('rc') files
if ($config->{rc}) {
        for my $rc (split("\0", $config->{rc})) {
                open(my $fh, $rc);
                $config = { %$config, %{ CGI->new($fh)->Vars } };
                close($fh);
        }
}

### Now, populate essential config if not specified already
$config->{user} //= getpwuid($<);
$config->{folder} //= 'Inbox';

### The IMAP client object that we will deal with all thru the script
my $client;

### For whatever reason, logout upon exit
END {
        $client->logout() or die "Could not logout: $@\n";
}

use Term::ReadKey;

sub collect_password 
{
        $config->{password} = eval {
                print "Enter Password for $config->{user} : \n";
                ReadMode 4; # Turn off controls keys
                my $p = ReadLine(0);
                ReadMode 0; # Reset tty mode before exiting
                chomp($p);
                return $p;
        }
}

### Collect password if not already known
if (!$config->{password} or $config->{password} eq '-') {
        collect_password();
}


sub connect_with_server()
{
        # Connect to IMAP server
        $client = Mail::IMAPClient->new(
                Server   => $config->{server},
                User     => $config->{user},
                Password => ($config->{password} =~ /^[01]*$/ ? pack('B*', $config->{password}) : $config->{password}),
                Port     => 993,
                Ssl      =>  1,
                IgnoreSizeErrors => 1,  ### Certain (caching) servers, like Exchange 2007, often report the wrong message size. 
                                        ### Instead of chopping the message into a size that it fits the specified size, 
                                        ### the reported size will be simply ignored when this parameter is set to 1.
        ) or die "ERROR: Cannot connect through IMAPClient: $!\n";

}

### Connect with IMAP server and initialize the client object
connect_with_server();


my $results; ### Current search results;

### Folder and Search be specified for a sumamry
sub print_summary()
{
        my $folder = $config->{folder} or die "no folder specified\n";
        my $search = $config->{search};

        $client->select($folder) or die "ERROR: Can not select folder=$folder " . $client->LastError . "\n";
        $results = $client->search($search) or do { print "can not search '$search' " . $client->LastError . "\n"; return; };

        print "Found ", scalar(@$results), " messages\n";

        if ($config->{summary_only}) {
                        ### No interest in the detail
                        return;
        }

        ### Print search results as a TSV

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

if ($config->{folder} and $config->{search}) {
        print_summary();
}

### Determine if any action needs to be done, once the search results are known
my $action = $config->{'action'} || '';

### For large sized data sets, determine a size of chunk that can be attempted at one time.
### For moving or deleting, IMAP can not deal with large sizes in one go.
my $batch_size = $config->{'batch_size'} || 1000;


### Intent is to move search results to another folder on the server
if ($action =~ /move_to_mail_folder:(.*)$/) {
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

        ### IMAP requires an "expunge" to FINALIZE the message movements.
        $client->expunge($config->{folder}) or die "Could not expung folder = $config->{folder}\n";
        exit(0);
}

### Intent is to copy the search results to a local folder on the computer.
### Essentially saves the MIME to local disk 
if ($action =~ /copy_to_local_folder:(.*)$/) {
        my $copy_to_local_folder = $1;
        my $folder = $config->{folder};
        for my $i (@$results) {
                $client->message_to_file("$copy_to_local_folder/$folder.$i.mime", $i) or die "ERROR: Can not copy_to_local_folder:$copy_to_local_folder " . $client->LastError . "\n";
                print "\t $i copied to local $copy_to_local_folder/$folder.$i.mime\n";
        }
        exit(0);
}

if ($action =~ /copy_to_local_file:(.*)$/) {
        my $copy_to_local_file = $1;
        $client->message_to_file($copy_to_local_file, @$results) or die "ERROR: Can not copy_to_local_file:$copy_to_local_file " . $client->LastError . "\n";
        print "\t Results copied to local file $copy_to_local_file\n";
        exit(0);
}

### DANGEROUS - Use it if and only if you understand what you are doing!
### Intent is to delete the search results. NO RECOVERY once deleted.
if ($action eq 'delete_message') {
        my $tmp;

        my $count = @$results;
        print "\nDeleting $count messages from folder=$config->{folder}\n";

        while (0 != @$results) {
                my $batch = [ splice(@$results, 0, $batch_size) ];
                my $batch_count = @$batch;
                my $tmp = $client->delete_message(@$batch);
                if (!defined($tmp)) {
                        die "ERROR: Can not delete $batch_count messages from $config->{folder} " . $client->LastError . "\n"; 
                }
                if ($tmp != $batch_count) {
                        warn "WARNING: Only $tmp of the $batch_count messages could be deleted\n";
                }
                print "\t $batch_count messages deleted from folder=$config->{folder}\n";
        }

        ### IMAP requires an "expunge" to FINALIZE the message movements.
        $client->expunge($config->{folder}) or die "Could not expung folder = $config->{folder}\n";
        print "$count messages deleted from folder=$config->{folder}\n";

        exit(0);
}



### Interactive shell to examine results
if ($action ne 'shell') {
        exit(0);
}

my $term = Term::ReadLine->new('IMAP Shell');
my $prompt = derive_prompt();

my $line;

my $body_cached = {};

while ( defined ($line = $term->readline($prompt)) ) {
        my $q1_hash_ref = CGI->new(join(';', map { s/'//g; s/"//g; $_ } ($line =~ /\S*$RE{quoted}\S*|\S+/g)))->Vars;

        $config =  { %$config, %$q1_hash_ref };


        if ($q1_hash_ref->{folder} or $q1_hash_ref->{search}) {
                print_summary();
                delete $config->{read};
        }

        if ($config->{folder} and !$config->{search}) {
                print "'search' not specified. You may specify search=ALL..\n";
                if ($client->selectable($config->{folder})) {
                        my $message_count = $client->message_count($config->{folder}) || 0;
                        print "\t* $config->{folder} ($message_count)\n";
                } else {
                        print "WARNING: folder=$config->{folder} is not selectable\n";
                }
        }

        if (!$config->{folder}) {
                print "'folder' not specified. Please specify one...\n";
                print "folders:\n";
                for my $i ($client->folders()) {
                        next if (!$client->selectable($i));
                        my $message_count = $client->message_count($i) || 0;
                        print "\t* $i (", $message_count, ")\n";
                }
        }

        if ($config->{quit}) {
                        exit(0);
        }

        if ($config->{read}) {
                $body_cached = {};
                my $msg_id = $config->{read};
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

                print "[ATTACHMENTS]\n\n";
                my $count = 1;
                for my $i ($stripper->attachments) {
                        print "$count) $i->{filename}, $i->{content_type}, ", length($i->{payload}), "\n"; ###filename, content_type, payload
                        $count++
                }
                print "[/ATTACHMENTS]\n\n";

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

                if ($config->{detail}) {
                        print $mime;
                }
                #delete $config->{read};
        }

        $prompt = derive_prompt();
}

sub derive_prompt
{
        #return CGI->new($current)->query_string;

        my $tmp = '';   
        $tmp = "[$config->{folder}" if ($config->{folder});
        if ($tmp) {
                if ($config->{search}) {
                        $tmp .= "|$config->{search}";
                        if ($config->{read}) {
                                $tmp .= "|$config->{read}";
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

        #my @parts = $in->parts(); 
        #my $num_parts = @parts;
        #my @ct_parts = map { $_->content_type } @parts;
        #my $ct_parts = join(',', @ct_parts);

        my @sub_parts = $in->subparts();
        my $num_subparts = @sub_parts;

        #my @ct_subparts = map { $_->content_type } @sub_parts;
        #my $ct_subparts = join(',', @ct_subparts);

        print "\t" x $depth, "$label) $ct\n"; 

        if ($config->{save} && 0 == $num_subparts) {
                ### If 'save' folder is specified, save all attachments (only leaf nodes)
                my $dirname = "$config->{save}/$config->{read}/";
                if (! -d $dirname) {
                        my $tmp = make_path($dirname);
                        if (!$tmp) {
                                die "ERROR: Can not make path '$dirname'\n";
                        }
                }
                my $filename = "$dirname/$label";
                if ($ct =~ /name="(.*)"/) {
                        $filename .= ".$1";
                } elsif ($ct =~ /(.*?)\//) {
                        $filename .= ".$1";
                }
                open(my $fh, ">$filename");
                binmode($fh);
                print $fh $in->body;
                close($fh);
        }

        if ($ct =~ /text\/plain/) {
                $body_cached->{$label} = $in->body;
        }

        my $tmp = 1;
        for my $p (@sub_parts) {
                traverse_mime($p, "$label.$tmp", $depth+1);
                $tmp++;
        }
}

