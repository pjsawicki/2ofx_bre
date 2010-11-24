#!/usr/bin/env perl -wT -CIO
#######################################################################
# 2OFX: Mark J Cox <mark@awe.com> 18 Nov 1999       www.awe.com/mark
#
# Convert the HTML pages giving your CAHOOT or EGG credit card statement
# into QIF or OFX, allowing import to things like Microsoft Money.
#
# Log into the EGG site and select either the statement, then select
# "printer friendly", then view source and save the file.  
# Run this program on the file:
#
#        perl 2ofx.pl < egg101199.html > egg101199.qif
# 
# If you are using CAHOOT and have more than one page, save each page
# and do:
#        perl 2ofx.pl page1.html page2.html > cahoot101199.qif
#
# Then import the file into your favourite program.  I'd recommend using
# MS Money 2001 as it will try to reconcile and match up entries in your
# account as well as marking the cleared entries as electronic ("E").
#
# Why don't they have a link to download a QIF version of your statement?
# This script will break if they change the format of their pages
#
# No warranty at all; Copy and use freely as long as this entire header
# section stays intact; send me your updates!
#######################################################################
# Version 1.00 18Nov99: first version
#         1.01 21Dec99: cope with payments
#         1.10 22Mar01: Deal with new format, quick hack to get it running
#         1.11 23Mar01: Ignore punctuation
#         2.01 24Mar01: Add Cahoot and merge in OFX support (use -o)
#         2.02 13May01: Add American Express, don't exceed money import lengths
#         2.03 04Jun01: Fix American Express card payments
#         2.04 21Dec02: New EGG support, mostly working
#         2.05 06Jan06: Handle a * in EGG payee name (paypal!)
#         2.06 16Jun06: Quick EGG Money support
#         2.07 09Jul06: Fix EGG normal statement (thanks Fitz for spotting)
#         2.08 17Nov06: Fix EGG Money statement (no more 'print' option)
#         2.09 29Nov06: Fix Cahoot statement (thanks Nick)
#         3.01 15Apr07: Robert Rasiewicz <robert.wk@gmail.com> - added multibank.pl (from print)
#         3.02 16May07: Robert Rasiewicz <robert.wk@gmail.com> - fixes for multibank.pl html variations
#         3.03 18Aug08: Tomasz Domin <zwirek at gmail.com> - initial version for mbank with all other formats temporary removed
#         4.01 05May09: Robert Rasiewicz <robert.wk@gmail.com> - added multibank csv support (ofx only and decimal point hardcoded)
#         4.02 24Aug10: Robert Rasiewicz <robert.wk@gmail.com> - multibank csv fix after column names change (may still be chopping the first letters)
#         4.03 11Sep10: Robert Rasiewicz <robert.wk@gmail.com> - not chopping first letters anymore, ofx memo field value fix
use strict;
use utf8;
use encoding 'utf8';
use Digest::MD5 qw(md5_hex);
use Encode qw (encode_utf8);

my @transactions;
my %s;

my $charset		= "utf-8";
my $language	= "POL";
my $acctype		= "CHECKING";
my $currency	= "PLN";

sub removegunk
{
    my ($line) = @_;
    $line =~ s/[\r\n]//g;  #Line endings
    $line =~ s/<[^>]+>//g; #HTML
    $line =~ s/&[^;]+;//g; #special characters
    $line =~ s/^\s+//g; #leading whitespace 
    $line =~ s/\s+$//g; #trailing whitespace
    return $line;
}

sub ofxoutput {
	my ($se,$mi,$ho,$d,$m,$y) = gmtime(time()); # can't assume strftime

	my $nowdate = sprintf("%04d%02d%02d%02d%02d%02d",$y+1900,$m+1,$d,$ho,$mi,$se);

	print <<EOH;
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:$charset
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
	<SIGNONMSGSRSV1>
		<SONRS>
			<STATUS>
				<CODE>0
				<SEVERITY>INFO
			</STATUS>
			<DTSERVER>$nowdate
			<LANGUAGE>$language
		</SONRS>
	</SIGNONMSGSRSV1>
	<BANKMSGSRSV1>
		<STMTTRNRS>
			<TRNUID>1
			<STATUS>
				<CODE>0
				<SEVERITY>INFO
			</STATUS>
			<STMTRS>
				<CURDEF>$currency
				<BANKACCTFROM>
					<BANKID>$s{'bankid'}
					<ACCTID>$s{'account'}
					<ACCTTYPE>$acctype
				</BANKACCTFROM>
				<BANKTRANLIST>
					<DTSTART>$s{'from'}
					<DTEND>$s{'to'}
EOH

	for (my$i=0;$i<$#transactions;$i++) {
		print "\t<STMTTRN>\n";

		print "\t\t<TRNTYPE>";
		if ($transactions[$i]{'type'}) {
			print $transactions[$i]{'type'};
		} else {
			print "DEBIT";
		}
		print "\n";

		print "\t\t<DTPOSTED>$transactions[$i]{'date'}\n";
		print "\t\t<FITID>$transactions[$i]{'tid'}\n";
		print "\t\t<TRNAMT>$transactions[$i]{'amount'}\n";
		print "\t\t<NAME>$transactions[$i]{'payee'}\n";
		print "\t\t<MEMO>$transactions[$i]{'memo'}\n";
		print "\t</STMTTRN>\n\n";
	}

	print "\t\t\t</BANKTRANLIST>\n";

	print <<EOH;
			</STMTRS>
		</STMTTRNRS>
	</BANKMSGSRSV1>
</OFX>
EOH
}

sub skip_lines {
    my ($lines)=@_;
    for (my $i=0; $i<($lines+1); $i++) {
    	$_ = <>;
    }
    return $_;
}

#sub mbank_karta_header {        
#    if (/Z RACHUNKU KARTY KREDYTOWEJ/) {
#        $s{'account'} = "Karta Kredytowa";
#	    $s{'bankid'} = "BREXPLPWMUL";
#    }
#
#    if (/(\d\d\d\d)-(\d\d)-(\d\d) DO (\d\d\d\d)-(\d\d)-(\d\d)/) {
#		 
#        $fy = $1;
#        $fm = $2;
#        $fd = $3;
#        
#        $ty = $4;
#        $tm = $5;
#        $td = $6;
#	    $s{'from'} = sprintf "%04d%02d%02d", $fy,$fm,$fd;
#	    $s{'to'} = sprintf "%04d%02d%02d", $ty,$tm,$td;
#    }
#    $language = "POL";
#    $acctype = "CHECKING";
#    $currency = "PLN";
#    
#   
#}
#sub mbank_karta_transactions
#{
##Nr oper.;#Data oper.;#Data ksiŕg.;#Rodzaj operacji;#Szczegˇ-y operacji;#Kwota w walucie oryginalnej;#Waluta oryginalna;#Kwota w PLN;
#	
#	if (m/(.*);(\d{4})-(\d{2})-(\d{2});\d{4}-\d{2}-\d{2};(.*);(.*);(.*);(.*);(.*);/) {
#	    $d = $4;
#	    $m = $3;
#	    $y = $2;
#	} else {
#		return;
#	}
#
#    # remove html variations
#	$memo = $5." ".$6;
#	$amount = $9;
#
#
#	$amount =~ s/[^0-9,-]//g;
#	$amount =~ s/,/\./g;
#	
#	if ($memo=~ /ODSETKI/)
#	{
#		$ttype="INT";
#	}
#	elsif ($memo=~ /PRZELEW WEWN/)
#	{
#		$ttype="CREDIT";
#	}
#	elsif ($memo=~ /ZAKUP PRZY/)
#	{
#		$ttype="DEBIT";
#	}
#
#	else
#	{
#		$ttype="DEBIT";
#	}
#
#	$payee=$memo;
#	$data=sprintf "%04d%02d%02d",$y,$m,$d;
#
#    $transactions[$transaction]{'date'} = $data;
#    $transactions[$transaction]{'memo'} = $memo ;
#    $transactions[$transaction]{'type'} = $ttype ;
#	$transactions[$transaction]{'payee'} = $payee;
#    $transactions[$transaction]{'amount'}=$amount;
#	$transactions[$transaction]{'tid'} = $date . "T".md5_hex($memo);
#    $transaction++;
#}
#sub mbank_pl_header
#{        
#    if (m/#Numer rachunku;/) {
#        $s{'account'} = skip_lines(0);
#		$s{'account'} =~ s/;//;
#		$s{'account'} =~ s/''//g;
#	    $s{'bankid'} = "BREXPLPWMUL";
#    }
#
#    if (m/#Za okres:;/) {
#    	$period = skip_lines(0);
#		$period =~ /(\d\d)\.(\d\d)\.(\d\d\d\d);(\d\d\).\(d\d)\.(\d\d\d\d);/;
#		 #Elektroniczne zestawienie operacji za okres od \s*(\d\d\d\d)-(\d\d)-(\d\d)\s*do\s*(\d\d\d\d)-(\d\d)-(\d\d)
#		 
#        $fy = $3;
#        $fm = $2;
#        $fd = $1;
#        
#        $ty = $6;
#        $tm = $5;
#        $td = $4;
#	    $s{'from'} = sprintf "%04d%02d%02d", $fy,$fm,$fd;
#	    $s{'to'} = sprintf "%04d%02d%02d", $ty,$tm,$td;
#	    
#	    #$charset = "iso-8859-2";
#	    $language = "POL";
#	    $acctype = "CHECKING";
#    }
#    
#    if (m/Waluta;/) {
#        $currency = removegunk(skip_lines(0));
#		$currency =~ s/;//;
#    }    
#}
#
#sub mbank_pl_transactions {
#	#Data operacji;#Data księgowania;#Opis operacji;#Nazwa;#Rachunek;#Nazwa Banku;#Opis dodatkowy;#Kwota;#Saldo po operacji
#	
#	if (m/(\d\d\d\d)-(\d\d)-(\d\d);\d\d\d\d-\d\d-\d\d;"(.*)";(.*);(.*);/) {
#	    $d = $3;
#	    $m = $2;
#	    $y = $1;
#	} else {
#		return;
#	}
#		
#    # remove html variations
#	$memo = $4;
#	$amount = $5;
#
#	$amount =~ s/[^0-9,-]//g;
#	$amount =~ s/,/\./g;
#	
#	if ($memo=~ /KAPITALIZACJA ODSETEK/) {
#		$ttype="INT";
#	} elsif ($memo=~ /WYPLATA W BANKOMACIE/) {
#		$ttype="ATM";
#	} elsif ($memo=~ /PODATEK OD ODSETEK KAPITALOWYCH/) {
#		$ttype="DEBIT";
#	} elsif ($memo=~ /KREDYT/) {
#		$ttype="DEBIT";
#	} elsif ($memo=~ /PRZELEW/) {
#		$ttype="XFER";
#	} elsif ($memo=~ /ZAKUP/) {
#		$ttype="XFER";
#		$payee=$memo;
#	} else {
#		$ttype="DEBIT";
#	}
#
#	$payee=$memo;
#	$data=sprintf "%04d%02d%02d",$y,$m,$d;
#
#    $transactions[$transaction]{'date'}=$data;
#    $transactions[$transaction]{'memo'}=$memo ;
#    $transactions[$transaction]{'type'}=$ttype ;
#	$transactions[$transaction]{'payee'} = $payee;
#    $transactions[$transaction]{'amount'}=$amount;
#	$transactions[$transaction]{'tid'} = $data."T".md5_hex($memo);
#    $transaction++;
#}

sub multibank_pl_header
{        
    if (m/Numer rachunku;/) {
        $s{'account'} = skip_lines(0);
		$s{'account'} =~ s/;//;
		$s{'account'} =~ s/'//g;
	    $s{'bankid'} = "BREXPLPWMUL";
    }

    if (m/za okres:;/) {
        my $period = skip_lines(0);
		#  od 2009-04-01;od 2009-04-30;
		#                ^^ blad w multibanku ..
		$period =~ /(\d\d)\.(\d\d)\.(\d\d\d\d);(\d\d\).\(d\d)\.(\d\d\d\d);/;
		
	    $s{'from'} = substr($period,3,4) . substr($period,8,2) . substr($period,11,2);
		$s{'to'} = substr($period,17,4) . substr($period,22,2) . substr($period,25,2);

    }
    
    if (m/Waluta;/) {
        $currency = removegunk(skip_lines(0));
		$currency =~ s/;//;
    }    
}

sub multibank_pl_transactions {	
	if (! /^\d{4}-\d{2}-\d{2}/) {
		return;
	}

	chomp;

	$_ =~ s/\s+/ /g;
	$_ =~ s/["']//g;

	my ($op_date, $acc_date, $type, $payee, $account, $title, $amount, $balance) = split(/;/, $_);

	my ($op_year, $op_month, $op_day) = split(/-/, $op_date);

	my $memo;

	if ($title) {
		$memo = "$type: $title";
	} else {
		$memo = "$title";
	}
		

	if ($payee =~ /^\s*$/) {
		if (length($memo) > 0) {
			$payee = $title;
		} else {
			$payee = $type;
		}
	}

	$amount =~ s/[^0-9,-]//g;
	$amount =~ s/,/\./g;

	my $ttype;
	
	if ($type=~ /KAPITALIZACJA ODSETEK/) {
		$ttype="INT";
	} elsif ($type=~ /WYPŁATA W BANKOMACIE/) {
		$ttype="ATM";
	} elsif ($type=~ /PODATEK OD ODSETEK KAPITAŁOWYCH/) {
		$ttype="DEBIT";
	} elsif ($type=~ /KREDYT/) {
		$ttype="DEBIT";
	} elsif ($type=~ /PRZELEW/) {
		$ttype="XFER";
	} elsif ($type=~ /ZAKUP/) {
		$ttype="XFER";
	} else {
		$ttype="DEBIT";
	}

	my $date = sprintf("%04d%02d%02d", $op_year, $op_month, $op_day);
	
	push (@transactions,
		{
			'amount'	=> $amount,
			'date'		=> $date,
			'memo'		=> $memo,
			'payee'		=> $payee,
			'tid'		=> $date . "T" . md5_hex(encode_utf8($memo)),
			'type'		=> $ttype,
		}
	);
}

#
# Main
#
#$mbank_pl=0;
#$mbank_karta=0;
my $multibank_pl=0;

my $filename = $ARGV[0];

while(<>) {
#    $mbank_pl = 1 if (/Elektroniczne zestawienie operacji/ && $mbank_pl ==0 && $multibank_pl==0);
#    $mbank_pl = 2 if (/#Data operacji;#Data księgowania;#Opis operacji;#Kwota;#Saldo po operacji;/ && $mbank_pl ==1);
#    mbank_pl_header() if ($mbank_pl == 1);
#    mbank_pl_transactions() if ($mbank_pl == 2);
#
#    $mbank_karta = 1 if (/Z RACHUNKU KARTY KREDYTOWEJ/ && $mbank_karta ==0);
#    $mbank_karta = 2 if (/#Nr oper.;#Data oper.;#Data księg.;#Rodzaj operacji;#Szczegóły operacji;#Kwota w walucie oryginalnej;#Waluta oryginalna;#Kwota w PLN;/ && $mbank_karta ==1);
#    mbank_karta_header() if ($mbank_karta == 1);
#    mbank_karta_transactions() if ($mbank_karta == 2);

    if (! $multibank_pl && $_ =~ /MultiBank/) {
    	$multibank_pl = 1;
    	next;
    }

    if ($multibank_pl == 1 && $_ =~ /Elektroniczne zestawienie operacji/) {
    	$multibank_pl = 2;
    	next;
   	}

    #if ($multibank_pl == 2 && $_ =~ /Data operacji;Data księgowania;Rodzaj transakcji;Dane odbiorcy/) {
    if ($multibank_pl == 2 && $_ =~ /Data operacji;Data ksi/) {
    	$multibank_pl = 3;
    	next;
   	}

    multibank_pl_header() if ($multibank_pl == 2);

    multibank_pl_transactions() if ($multibank_pl == 3);
}

#die "Didn't find anything that looked like a statement" 
#		 unless $multibank_pl or $mbank_pl or $mbank_karta;

if ($#transactions > 0) {
	ofxoutput();
	exit(0);
} else { 
	die "Couldn't find any transactions";
	exit 0;
}
