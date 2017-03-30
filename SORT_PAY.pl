use strict; use warnings;
use Spreadsheet::ParseExcel;
#######################################
# SORT_PAY - categorize payment sources
# SETUP ###############################
use constant { ICOL => 0, PCOL => 6, }; # Dependent on target file
my ($l, $book) = @ARGV;
open(my $log, '>>', $l) or die "ARG1 ERROR NEED LOGFILE";
my $parser = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse($book);
my $worksheet = $workbook->worksheet(0);
my ($row_min, $lastrow) = $worksheet->row_range();
my @payment; # FINAL EXTRACTON ARRAY
# GET BATCH NUMBER (4th part of final extraction string)
my $batch_cell = $worksheet->get_cell(0, 5);
my $batch_value = $batch_cell->value();
$batch_value =~ s/Batch number //;
# POPULATE ARRAYS - HASH (3rd element of final extraction string)
my @ttl = ("Payment - Check Total:", "Payment - Other Total:", "Payment - Cash Total:", "Payment - Visa Total:", "Payment - MasterCard Total:", "Payment - Discover Total:", "Payment - American Express Total:", "Payment - Other Card Total:");
my %pay = map{$_ => undef} @ttl;
for my $row (0..$lastrow) {
    my $icell = $worksheet->get_cell($row, ICOL);
    my $ivalue = $icell->value();
    next unless (exists $pay{$ivalue});
    my $pay_cell = $worksheet->get_cell($row, PCOL);
    my $ttl_pay = $pay_cell->value();
    $ttl_pay=~ s/\$//; $ttl_pay =~ s/\(//; $ttl_pay =~ s/\)//; $ttl_pay =~ s/\,//;
    next if ($ttl_pay eq "" or $ttl_pay eq "0.00");
    my $amt = $ttl_pay;
    my $last_loop = $ivalue; $last_loop =~ s/ Total://;
    my $loop_row =  $row;
    my $success_value = 0; my $success_cell = 0;
    until ($success_value eq $last_loop) {
        $loop_row--;
        $success_cell = $worksheet->get_cell($loop_row, ICOL);
        $success_value = $success_cell->value(); my $name = $success_value; $name =~ s/ /_/;
        my $amtcell = $worksheet->get_cell($loop_row, PCOL);
        $amt = $amtcell->value();
        $amt=~ s/\$//; $amt =~ s/\(//; $amt =~ s/\)//; $amt =~ s/\,//;
        next if ($amt eq "" or $amt eq "0.00");
        my $type = $last_loop; $type =~ s/.* //;
        my $prepay = "$name $amt $type $batch_value";
        push(@payment, $prepay);
    }
}
for (@payment) { print "$_\n"; print $log "$_\n"; }
