use strict; use warnings;
use Spreadsheet::ParseXLSX;
#########################################
# DIS - print values of spreadsheet range
### DIS.pl start_sheet end_sheet row_1 col_1 (of cell, additional cell optional)
### DIS.pl sheet1 sheet2 row_1 col_1 row_2 col_2
# row/cell start at 0
# in Excel set A1 = 0; then in B2 set =A1+1; (this will help)
# SETUP #####################################
my ($book, $start, $end, $row1, $col1, $row2, $col2) = @ARGV;
die ("no book arg; book start end cell1 cell2 cell3") unless defined $book;
die ("no start arg; book start end cell1 cell2 cell3") unless defined $start;
die ("no end arg; book start end cell1 cell2 cell3") unless defined $end;
die ("no cell arg; book start end cell1 cell2 cell3") unless defined $row1;
open(my $bfh, '<', $book) or die ("FAIL OPEN $book\n");
my $parser = Spreadsheet::ParseXLSX->new;
my $wb = $parser->parse($book);
die ("parser error") unless defined $wb;
# SHEETS ###################################
my @sheets;
foreach ($wb->worksheets())
    { push @sheets, $_->get_name(); }
my $s = 0; my $e = 0;
foreach (@sheets)
    { last if $_ eq $start; $s++; }
foreach (@sheets)
    { last if $_ eq $end; $e++; }
# GET DATA ##################################
while ($s <= $e) {
    my $ws = $wb->worksheet($s);
    my $cell1 = $ws->get_cell($row1, $col1);
    my $cell2 = $ws->get_cell($row2, $col2);
    $s++;
    my $sheet = $ws->get_name;
    my $value1 = $cell1->value;
# CELL 2 OPTIONAL ###########################
    my $value2 = $cell2->value if defined $cell2;
    print "$sheet : $value1";
    if (defined $value2)
        { print " : $value2"; }
    print "\n";
}
