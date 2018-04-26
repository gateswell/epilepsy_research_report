#the script is used for generating the epilepsy trios research result in excel format, also arranged every cell under doctor's suggestions addtionally.
#Usage: perl family-screen-candidate-auto_version5.pl -I <noInterGene.ComHet.family.refined.txt> -o <Research_sampleID_trios.xlsx> -r <epi_1046_genes.info_csh.txt>
#Author: Wei Ye, Shuhuan Cao
#Date: Thu Apr 26 11:15:13 CST 2018
#Version: 5
use Excel::Writer::XLSX;
use Encode;
use Getopt::Std;

my %opts;
getopts("I:o:r:h",\%opts);

my $input = $opts{I};	#input
my $output = $opts{o};
my $ref_in = $opts{r};

&print_usage unless (defined($input));
&print_usage unless (defined($output));
&print_usage unless (defined($ref_in));
&print_usage if (defined($opts{h}));

#open OUT, $ARGV[0];
open OUT, $input;
chomp ($header = <OUT>);
$header =~ s/#//g;
@header = split /\t/,$header;

unshift @header, "Repeat";
push @header, "Important";
#my $outworkbook = Excel::Writer::XLSX->new($ARGV[1]);
my $outworkbook = Excel::Writer::XLSX->new($output);
#my $fgrey = $outworkbook->add_format(
#	color	=> 'grey',
#);
#my $fboldgrey = $outworkbook->add_format(
#	bold	=> 1,
#	color	=> 'grey',
#);
my @unsort_header = @header;
shift @unsort_header;
unshift @unsort_header,'Inheritance';
unshift @unsort_header,'GeneCategories';
unshift @unsort_header,'Repeat';
my %index;
my $num_index = 0;
for(@unsort_header){
	$index{$_} = $num_index;
	$num_index ++;
}
#print "@unsort_header";
my @format;
$format[1] = $outworkbook -> add_format();	#A
$format[1] -> set_bold(1);
$format[1] -> set_font('Times New Roman');
$format[1] -> set_align( 'left' );

$format[2] = $outworkbook -> add_format();	#B
$format[2] -> set_bold(1);
$format[2] -> set_font('Times New Roman');
$format[2] -> set_align( 'left' );

$format[3] = $outworkbook -> add_format();	#C
$format[3] -> set_bold(1);
$format[3] -> set_font('Times New Roman');
$format[3] -> set_align( 'left' );

$format[4] = $outworkbook -> add_format();	#D
$format[4] -> set_bold(1);
$format[4] -> set_font('Times New Roman');
$format[4] -> set_align( 'left' );

$format[5] = $outworkbook -> add_format();	#E
$format[5] -> set_bold(1);
$format[5] -> set_font('Times New Roman');
$format[5] -> set_align( 'center' );

$format[6] = $outworkbook -> add_format();	#F
$format[6] -> set_bold(1);
$format[6] -> set_font('Times New Roman');
$format[6] -> set_align( 'center' );

$format[7] = $outworkbook -> add_format();	#G
$format[7] -> set_bold(1);
$format[7] -> set_font('Times New Roman');
$format[7] -> set_align( 'left' );

$format[8] = $outworkbook -> add_format();	#H
$format[8] -> set_bold(1);
$format[8] -> set_font('Times New Roman');
$format[8] -> set_align( 'right' );

$format[9] = $outworkbook -> add_format();	#I
$format[9] -> set_bold(1);
$format[9] -> set_font('Times New Roman');
$format[9] -> set_align( 'right' );

$format[10] = $outworkbook -> add_format();	#J
$format[10] -> set_bold(1);
$format[10] -> set_font('Times New Roman');
$format[10] -> set_align( 'center' );

$format[11] = $outworkbook -> add_format();	#K
$format[11] -> set_bold(1);
$format[11] -> set_font('Times New Roman');
$format[11] -> set_align( 'center' );

$format[12] = $outworkbook -> add_format();	#L
$format[12] -> set_bold(1);
$format[12] -> set_font('Times New Roman');
$format[12] -> set_align( 'center' );

$format[13] = $outworkbook -> add_format();	#M
$format[13] -> set_bold(1);
$format[13] -> set_font('Times New Roman');
$format[13] -> set_align( 'right' );

$format[14] = $outworkbook -> add_format();	#N
$format[14] -> set_bold(1);
$format[14] -> set_font('Times New Roman');
$format[14] -> set_align( 'left' );

$format[15] = $outworkbook -> add_format();	#O
$format[15] -> set_bold(1);
$format[15] -> set_font('Times New Roman');
$format[15] -> set_align( 'right' );

$format[16] = $outworkbook -> add_format();	#P
$format[16] -> set_bold(1);
$format[16] -> set_font('Times New Roman');
$format[16] -> set_align( 'right' );

$format[17] = $outworkbook -> add_format();	#Q
$format[17] -> set_bold(1);
$format[17] -> set_font('Times New Roman');
$format[17] -> set_align( 'right' );

$format[18] = $outworkbook -> add_format();	#R
$format[18] -> set_bold(1);
$format[18] -> set_font('Times New Roman');
$format[18] -> set_align( 'right' );

$format[19] = $outworkbook -> add_format();	#S
$format[19] -> set_bold(1);
$format[19] -> set_font('Times New Roman');
$format[19] -> set_align( 'right' );

$format[20] = $outworkbook -> add_format();	#T
$format[20] -> set_bold(1);
$format[20] -> set_font('Times New Roman');
$format[20] -> set_align( 'right' );

$format[21] = $outworkbook -> add_format();	#U
$format[21] -> set_bold(1);
$format[21] -> set_font('Times New Roman');
$format[21] -> set_align( 'left' );

$format[22] = $outworkbook -> add_format();	#V
$format[22] -> set_bold(1);
$format[22] -> set_font('Times New Roman');
$format[22] -> set_align( 'left' );

$format[23] = $outworkbook -> add_format();	#W
$format[23] -> set_bold(1);
$format[23] -> set_font('Times New Roman');
$format[23] -> set_align( 'left' );

$format[24] = $outworkbook -> add_format();	#X
$format[24] -> set_bold(1);
$format[24] -> set_font('Times New Roman');
$format[24] -> set_align( 'left' );

$format[25] = $outworkbook -> add_format();	#Y
$format[25] -> set_bold(1);
$format[25] -> set_font('Times New Roman');
$format[25] -> set_align( 'left' );

$format[26] = $outworkbook -> add_format();	#Z
$format[26] -> set_bold(1);
$format[26] -> set_font('Times New Roman');
$format[26] -> set_align( 'left' );

$format[27] = $outworkbook -> add_format();	#AA
$format[27] -> set_bold(1);
$format[27] -> set_font('Times New Roman');
$format[27] -> set_align( 'left' );

my @format2;									#内容格式2
$format2[1] = $outworkbook -> add_format();	#A
$format2[1] -> set_font('Times New Roman');
$format2[1] -> set_align( 'left' );

$format2[2] = $outworkbook -> add_format();	#B
$format2[2] -> set_font('Times New Roman');
$format2[2] -> set_align( 'center' );

$format2[3] = $outworkbook -> add_format();	#C
$format2[3] -> set_font('Times New Roman');
$format2[3] -> set_align( 'left' );

$format2[4] = $outworkbook -> add_format();	#D
$format2[4] -> set_font('Times New Roman');
$format2[4] -> set_align( 'center' );

$format2[5] = $outworkbook -> add_format();	#E
$format2[5] -> set_font('Times New Roman');
$format2[5] -> set_align( 'center' );

$format2[6] = $outworkbook -> add_format();	#F
$format2[6] -> set_font('Times New Roman');
$format2[6] -> set_align( 'left' );

$format2[7] = $outworkbook -> add_format();	#G
$format2[7] -> set_font('Times New Roman');
$format2[7] -> set_align( 'left' );

$format2[8] = $outworkbook -> add_format();	#H
$format2[8] -> set_font('Times New Roman');
$format2[8] -> set_align( 'right' );

$format2[9] = $outworkbook -> add_format();	#I
$format2[9] -> set_font('Times New Roman');
$format2[9] -> set_align( 'right' );

$format2[10] = $outworkbook -> add_format();	#J
$format2[10] -> set_font('Times New Roman');
$format2[10] -> set_align( 'center' );

$format2[11] = $outworkbook -> add_format();	#K
$format2[11] -> set_font('Times New Roman');
$format2[11] -> set_align( 'center' );

$format2[12] = $outworkbook -> add_format();	#L
$format2[12] -> set_font('Times New Roman');
$format2[12] -> set_align( 'center' );

$format2[13] = $outworkbook -> add_format();	#M
$format2[13] -> set_font('Times New Roman');
$format2[13] -> set_align( 'center' );

$format2[14] = $outworkbook -> add_format();	#N
$format2[14] -> set_font('Times New Roman');
$format2[14] -> set_align( 'center' );

$format2[15] = $outworkbook -> add_format();	#O
$format2[15] -> set_font('Times New Roman');
$format2[15] -> set_align( 'center' );

$format2[16] = $outworkbook -> add_format();	#P
$format2[16] -> set_font('Times New Roman');
$format2[16] -> set_align( 'center' );

$format2[17] = $outworkbook -> add_format();	#Q
$format2[17] -> set_font('Times New Roman');
$format2[17] -> set_align( 'left' );

$format2[18] = $outworkbook -> add_format();	#R
$format2[18] -> set_font('Times New Roman');
$format2[18] -> set_align( 'left' );

$format2[19] = $outworkbook -> add_format();	#S
$format2[19] -> set_font('Times New Roman');
$format2[19] -> set_align( 'left' );

$format2[20] = $outworkbook -> add_format();	#T
$format2[20] -> set_font('Times New Roman');
$format2[20] -> set_align( 'center' );

$format2[21] = $outworkbook -> add_format();	#U
$format2[21] -> set_font('Times New Roman');
$format2[21] -> set_align( 'left' );

$format2[22] = $outworkbook -> add_format();	#V
$format2[22] -> set_font('Times New Roman');
$format2[22] -> set_align( 'left' );

$format2[23] = $outworkbook -> add_format();	#W
$format2[23] -> set_font('Times New Roman');
$format2[23] -> set_align( 'left' );

$format2[24] = $outworkbook -> add_format();	#X
$format2[24] -> set_font('Times New Roman');
$format2[24] -> set_align( 'left' );

$format2[25] = $outworkbook -> add_format();	#Y
$format2[25] -> set_font('Times New Roman');
$format2[25] -> set_align( 'left' );

$format2[26] = $outworkbook -> add_format();	#Z
$format2[26] -> set_font('Times New Roman');
$format2[26] -> set_align( 'left' );

$format2[27] = $outworkbook -> add_format();	#AA
$format2[27] -> set_font('Times New Roman');
$format2[27] -> set_align( 'left' );

my @format3;									#内容格式3
$format3[1] = $outworkbook -> add_format();	#A
$format3[1] -> set_color('grey');
$format3[1] -> set_font('Times New Roman');
$format3[1] -> set_align( 'left' );

$format3[2] = $outworkbook -> add_format();	#B
$format3[2] -> set_color('grey');
$format3[2] -> set_font('Times New Roman');
$format3[2] -> set_align( 'center' );

$format3[3] = $outworkbook -> add_format();	#C
$format3[3] -> set_color('grey');
$format3[3] -> set_font('Times New Roman');
$format3[3] -> set_align( 'left' );

$format3[4] = $outworkbook -> add_format();	#D
$format3[4] -> set_color('grey');
$format3[4] -> set_font('Times New Roman');
$format3[4] -> set_align( 'center' );

$format3[5] = $outworkbook -> add_format();	#E
$format3[5] -> set_color('grey');
$format3[5] -> set_font('Times New Roman');
$format3[5] -> set_align( 'center' );

$format3[6] = $outworkbook -> add_format();	#F
$format3[6] -> set_color('grey');
$format3[6] -> set_font('Times New Roman');
$format3[6] -> set_align( 'left' );

$format3[7] = $outworkbook -> add_format();	#G
$format3[7] -> set_color('grey');
$format3[7] -> set_font('Times New Roman');
$format3[7] -> set_align( 'left' );

$format3[8] = $outworkbook -> add_format();	#H
$format3[8] -> set_color('grey');
$format3[8] -> set_font('Times New Roman');
$format3[8] -> set_align( 'right' );

$format3[9] = $outworkbook -> add_format();	#I
$format3[9] -> set_color('grey');
$format3[9] -> set_font('Times New Roman');
$format3[9] -> set_align( 'right' );

$format3[10] = $outworkbook -> add_format();	#J
$format3[10] -> set_color('grey');
$format3[10] -> set_font('Times New Roman');
$format3[10] -> set_align( 'center' );

$format3[11] = $outworkbook -> add_format();	#K
$format3[11] -> set_color('grey');
$format3[11] -> set_font('Times New Roman');
$format3[11] -> set_align( 'center' );

$format3[12] = $outworkbook -> add_format();	#L
$format3[12] -> set_color('grey');
$format3[12] -> set_font('Times New Roman');
$format3[12] -> set_align( 'center' );

$format3[13] = $outworkbook -> add_format();	#M
$format3[13] -> set_color('grey');
$format3[13] -> set_font('Times New Roman');
$format3[13] -> set_align( 'right' );

$format3[14] = $outworkbook -> add_format();	#N
$format3[14] -> set_color('grey');
$format3[14] -> set_font('Times New Roman');
$format3[14] -> set_align( 'center' );

$format3[15] = $outworkbook -> add_format();	#O
$format3[15] -> set_color('grey');
$format3[15] -> set_font('Times New Roman');
$format3[15] -> set_align( 'center' );

$format3[16] = $outworkbook -> add_format();	#P
$format3[16] -> set_color('grey');
$format3[16] -> set_font('Times New Roman');
$format3[16] -> set_align( 'center' );

$format3[17] = $outworkbook -> add_format();	#Q
$format3[17] -> set_color('grey');
$format3[17] -> set_font('Times New Roman');
$format3[17] -> set_align( 'left' );

$format3[18] = $outworkbook -> add_format();	#R
$format3[18] -> set_color('grey');
$format3[18] -> set_font('Times New Roman');
$format3[18] -> set_align( 'left' );

$format3[19] = $outworkbook -> add_format();	#S
$format3[19] -> set_color('grey');
$format3[19] -> set_font('Times New Roman');
$format3[19] -> set_align( 'left' );

$format3[20] = $outworkbook -> add_format();	#T
$format3[20] -> set_color('grey');
$format3[20] -> set_font('Times New Roman');
$format3[20] -> set_align( 'center' );

$format3[21] = $outworkbook -> add_format();	#U
$format3[21] -> set_color('grey');
$format3[21] -> set_font('Times New Roman');
$format3[21] -> set_align( 'left' );

$format3[22] = $outworkbook -> add_format();	#V
$format3[22] -> set_color('grey');
$format3[22] -> set_font('Times New Roman');
$format3[22] -> set_align( 'left' );

$format3[23] = $outworkbook -> add_format();	#W
$format3[23] -> set_color('grey');
$format3[23] -> set_font('Times New Roman');
$format3[23] -> set_align( 'left' );

$format3[24] = $outworkbook -> add_format();	#X
$format3[24] -> set_color('grey');
$format3[24] -> set_font('Times New Roman');
$format3[24] -> set_align( 'left' );

$format3[25] = $outworkbook -> add_format();	#Y
$format3[25] -> set_color('grey');
$format3[25] -> set_font('Times New Roman');
$format3[25] -> set_align( 'left' );

$format3[26] = $outworkbook -> add_format();	#Z
$format3[26] -> set_color('grey');
$format3[26] -> set_font('Times New Roman');
$format3[26] -> set_align( 'left' );

$format3[27] = $outworkbook -> add_format();	#AA
$format3[27] -> set_color('grey');
$format3[27] -> set_font('Times New Roman');
$format3[27] -> set_align( 'left' );

my @format4;									#内容格式4
$format4[1] = $outworkbook -> add_format();	#A
$format4[1] -> set_color('grey');
$format4[1] -> set_bold(1);
$format4[1] -> set_font('Times New Roman');
$format4[1] -> set_align( 'left' );

$format4[2] = $outworkbook -> add_format();	#B
$format4[2] -> set_color('grey');
$format4[2] -> set_bold(1);
$format4[2] -> set_font('Times New Roman');
$format4[2] -> set_align( 'center' );

$format4[3] = $outworkbook -> add_format();	#C
$format4[3] -> set_color('grey');
$format4[3] -> set_bold(1);
$format4[3] -> set_font('Times New Roman');
$format4[3] -> set_align( 'left' );

$format4[4] = $outworkbook -> add_format();	#D
$format4[4] -> set_color('grey');
$format4[4] -> set_bold(1);
$format4[4] -> set_font('Times New Roman');
$format4[4] -> set_align( 'center' );

$format4[5] = $outworkbook -> add_format();	#E
$format4[5] -> set_color('grey');
$format4[5] -> set_bold(1);
$format4[5] -> set_font('Times New Roman');
$format4[5] -> set_align( 'center' );

$format4[6] = $outworkbook -> add_format();	#F
$format4[6] -> set_color('grey');
$format4[6] -> set_bold(1);
$format4[6] -> set_font('Times New Roman');
$format4[6] -> set_align( 'left' );

$format4[7] = $outworkbook -> add_format();	#G
$format4[7] -> set_color('grey');
$format4[7] -> set_bold(1);
$format4[7] -> set_font('Times New Roman');
$format4[7] -> set_align( 'left' );

$format4[8] = $outworkbook -> add_format();	#H
$format4[8] -> set_color('grey');
$format4[8] -> set_bold(1);
$format4[8] -> set_font('Times New Roman');
$format4[8] -> set_align( 'right' );

$format4[9] = $outworkbook -> add_format();	#I
$format4[9] -> set_color('grey');
$format4[9] -> set_bold(1);
$format4[9] -> set_font('Times New Roman');
$format4[9] -> set_align( 'right' );

$format4[10] = $outworkbook -> add_format();	#J
$format4[10] -> set_color('grey');
$format4[10] -> set_bold(1);
$format4[10] -> set_font('Times New Roman');
$format4[10] -> set_align( 'center' );

$format4[11] = $outworkbook -> add_format();	#K
$format4[11] -> set_color('grey');
$format4[11] -> set_bold(1);
$format4[11] -> set_font('Times New Roman');
$format4[11] -> set_align( 'center' );

$format4[12] = $outworkbook -> add_format();	#L
$format4[12] -> set_color('grey');
$format4[12] -> set_bold(1);
$format4[12] -> set_font('Times New Roman');
$format4[12] -> set_align( 'center' );

$format4[13] = $outworkbook -> add_format();	#M
$format4[13] -> set_color('grey');
$format4[13] -> set_bold(1);
$format4[13] -> set_font('Times New Roman');
$format4[13] -> set_align( 'right' );

$format4[14] = $outworkbook -> add_format();	#N
$format4[14] -> set_color('grey');
$format4[14] -> set_bold(1);
$format4[14] -> set_font('Times New Roman');
$format4[14] -> set_align( 'center' );

$format4[15] = $outworkbook -> add_format();	#O
$format4[15] -> set_color('grey');
$format4[15] -> set_bold(1);
$format4[15] -> set_font('Times New Roman');
$format4[15] -> set_align( 'center' );

$format4[16] = $outworkbook -> add_format();	#P
$format4[16] -> set_color('grey');
$format4[16] -> set_bold(1);
$format4[16] -> set_font('Times New Roman');
$format4[16] -> set_align( 'center' );

$format4[17] = $outworkbook -> add_format();	#Q
$format4[17] -> set_color('grey');
$format4[17] -> set_bold(1);
$format4[17] -> set_font('Times New Roman');
$format4[17] -> set_align( 'left' );

$format4[18] = $outworkbook -> add_format();	#R
$format4[18] -> set_color('grey');
$format4[18] -> set_bold(1);
$format4[18] -> set_font('Times New Roman');
$format4[18] -> set_align( 'left' );

$format4[19] = $outworkbook -> add_format();	#S
$format4[19] -> set_color('grey');
$format4[19] -> set_bold(1);
$format4[19] -> set_font('Times New Roman');
$format4[19] -> set_align( 'left' );

$format4[20] = $outworkbook -> add_format();	#T
$format4[20] -> set_color('grey');
$format4[20] -> set_bold(1);
$format4[20] -> set_font('Times New Roman');
$format4[20] -> set_align( 'center' );

$format4[21] = $outworkbook -> add_format();	#U
$format4[21] -> set_color('grey');
$format4[21] -> set_bold(1);
$format4[21] -> set_font('Times New Roman');
$format4[21] -> set_align( 'left' );

$format4[22] = $outworkbook -> add_format();	#V
$format4[22] -> set_color('grey');
$format4[22] -> set_bold(1);
$format4[22] -> set_font('Times New Roman');
$format4[22] -> set_align( 'left' );

$format4[23] = $outworkbook -> add_format();	#W
$format4[23] -> set_color('grey');
$format4[23] -> set_bold(1);
$format4[23] -> set_font('Times New Roman');
$format4[23] -> set_align( 'left' );

$format4[24] = $outworkbook -> add_format();	#X
$format4[24] -> set_color('grey');
$format4[24] -> set_bold(1);
$format4[24] -> set_font('Times New Roman');
$format4[24] -> set_align( 'left' );

$format4[25] = $outworkbook -> add_format();	#Y
$format4[25] -> set_color('grey');
$format4[25] -> set_bold(1);
$format4[25] -> set_font('Times New Roman');
$format4[25] -> set_align( 'left' );

$format4[26] = $outworkbook -> add_format();	#Z
$format4[26] -> set_color('grey');
$format4[26] -> set_bold(1);
$format4[26] -> set_font('Times New Roman');
$format4[26] -> set_align( 'left' );

$format4[27] = $outworkbook -> add_format();	#AA
$format4[27] -> set_color('grey');
$format4[27] -> set_bold(1);
$format4[27] -> set_font('Times New Roman');
$format4[27] -> set_align( 'left' );

my @ordered_header = qw/Gene Repeat GeneCategories Inheritance CHR POS Consequence HGVSc HGVSp KG ESP ExAC Num_of_Harm Origin GT_Case GT_ControlF GT_ControlM Detail_Case Detail_ControlF Detail_ControlM avsnp150 REF ALT FREQ IMPACT ComHet Important/;
$WSOUT_Cdd = $outworkbook ->add_worksheet(decode("GB2312","KG和ExAC<0.005过滤"));
$WSOUT_Cdd -> freeze_panes(0,4);
$WSOUT_Cdd -> set_column('A:A',8.25);
$WSOUT_Cdd -> set_column('B:B',3.25);
$WSOUT_Cdd -> set_column('C:C',11.88);
$WSOUT_Cdd -> set_column('D:D',4.5);
$WSOUT_Cdd -> set_column('E:E',4.25);
$WSOUT_Cdd -> set_column('F:F',9.88);
$WSOUT_Cdd -> set_column('G:G',8.38);
$WSOUT_Cdd -> set_column('H:H',9.0);
$WSOUT_Cdd -> set_column('I:I',9.63);
$WSOUT_Cdd -> set_column('J:J',8.5);
$WSOUT_Cdd -> set_column('K:K',8.5);
$WSOUT_Cdd -> set_column('L:L',8.5);
$WSOUT_Cdd -> set_column('M:M',4.25);
$WSOUT_Cdd -> set_column('N:N',3.38);
$WSOUT_Cdd -> set_column('O:O',4.38);
$WSOUT_Cdd -> set_column('P:P',4.38);
$WSOUT_Cdd -> set_column('Q:Q',4.38);
$WSOUT_Cdd -> set_column('R:R',8.38);
$WSOUT_Cdd -> set_column('S:S',8.38);
$WSOUT_Cdd -> set_column('T:T',8.38);
$WSOUT_Cdd -> set_column('U:U',9.13);
$WSOUT_Cdd -> set_column('V:V',3.38);
$WSOUT_Cdd -> set_column('W:W',3.38);
$WSOUT_Cdd -> set_column('X:X',5.88);
$WSOUT_Cdd -> set_column('Y:Y',8.38);
$WSOUT_Cdd -> set_column('Z:Z',5.88);
$WSOUT_Cdd -> set_column('AA:AA',3.38);

foreach $i(0 .. $#ordered_header){
	$WSOUT_Cdd->write(0, $i, $ordered_header[$i],$format[$i+1]);
}

$WSOUT_Known = $outworkbook ->add_worksheet(decode("GB2312","Known过滤"));
$WSOUT_Known -> freeze_panes(0,4);
$WSOUT_Known -> set_column('A:A',8.25);
$WSOUT_Known -> set_column('B:B',3.25);
$WSOUT_Known -> set_column('C:C',11.88);
$WSOUT_Known -> set_column('D:D',4.5);
$WSOUT_Known -> set_column('E:E',4.25);
$WSOUT_Known -> set_column('F:F',9.88);
$WSOUT_Known -> set_column('G:G',8.38);
$WSOUT_Known -> set_column('H:H',9.0);
$WSOUT_Known -> set_column('I:I',9.63);
$WSOUT_Known -> set_column('J:J',8.5);
$WSOUT_Known -> set_column('K:K',8.5);
$WSOUT_Known -> set_column('L:L',8.5);
$WSOUT_Known -> set_column('M:M',4.25);
$WSOUT_Known -> set_column('N:N',3.38);
$WSOUT_Known -> set_column('O:O',4.38);
$WSOUT_Known -> set_column('P:P',4.38);
$WSOUT_Known -> set_column('Q:Q',4.38);
$WSOUT_Known -> set_column('R:R',8.38);
$WSOUT_Known -> set_column('S:S',8.38);
$WSOUT_Known -> set_column('T:T',8.38);
$WSOUT_Known -> set_column('U:U',9.13);
$WSOUT_Known -> set_column('V:V',3.38);
$WSOUT_Known -> set_column('W:W',3.38);
$WSOUT_Known -> set_column('X:X',5.88);
$WSOUT_Known -> set_column('Y:Y',8.38);
$WSOUT_Known -> set_column('Z:Z',5.88);
$WSOUT_Known -> set_column('AA:AA',3.38);
foreach $i(0 .. $#ordered_header){
	$WSOUT_Known->write(0, $i, $ordered_header[$i],$format[$i+1]);
}

my $WSOUT_Denovo = $outworkbook ->add_worksheet(decode("GB2312","DeNovo过滤"));
$WSOUT_Denovo -> freeze_panes(0,4);
$WSOUT_Denovo -> set_column('A:A',8.25);
$WSOUT_Denovo -> set_column('B:B',3.25);
$WSOUT_Denovo -> set_column('C:C',11.88);
$WSOUT_Denovo -> set_column('D:D',4.5);
$WSOUT_Denovo -> set_column('E:E',4.25);
$WSOUT_Denovo -> set_column('F:F',9.88);
$WSOUT_Denovo -> set_column('G:G',8.38);
$WSOUT_Denovo -> set_column('H:H',9.0);
$WSOUT_Denovo -> set_column('I:I',9.63);
$WSOUT_Denovo -> set_column('J:J',8.5);
$WSOUT_Denovo -> set_column('K:K',8.5);
$WSOUT_Denovo -> set_column('L:L',8.5);
$WSOUT_Denovo -> set_column('M:M',4.25);
$WSOUT_Denovo -> set_column('N:N',3.38);
$WSOUT_Denovo -> set_column('O:O',4.38);
$WSOUT_Denovo -> set_column('P:P',4.38);
$WSOUT_Denovo -> set_column('Q:Q',4.38);
$WSOUT_Denovo -> set_column('R:R',8.38);
$WSOUT_Denovo -> set_column('S:S',8.38);
$WSOUT_Denovo -> set_column('T:T',8.38);
$WSOUT_Denovo -> set_column('U:U',9.13);
$WSOUT_Denovo -> set_column('V:V',3.38);
$WSOUT_Denovo -> set_column('W:W',3.38);
$WSOUT_Denovo -> set_column('X:X',5.88);
$WSOUT_Denovo -> set_column('Y:Y',8.38);
$WSOUT_Denovo -> set_column('Z:Z',5.88);
$WSOUT_Denovo -> set_column('AA:AA',3.38);
foreach $i(0 .. $#ordered_header){
	$WSOUT_Denovo->write(0, $i, $ordered_header[$i],$format[$i+1]);
}

my $WSOUT_Domi = $outworkbook ->add_worksheet(decode("GB2312","Dominant过滤"));
$WSOUT_Domi -> freeze_panes(0,4);
$WSOUT_Domi -> set_column('A:A',8.25);
$WSOUT_Domi -> set_column('B:B',3.25);
$WSOUT_Domi -> set_column('C:C',11.88);
$WSOUT_Domi -> set_column('D:D',4.5);
$WSOUT_Domi -> set_column('E:E',4.25);
$WSOUT_Domi -> set_column('F:F',9.88);
$WSOUT_Domi -> set_column('G:G',8.38);
$WSOUT_Domi -> set_column('H:H',9.0);
$WSOUT_Domi -> set_column('I:I',9.63);
$WSOUT_Domi -> set_column('J:J',8.5);
$WSOUT_Domi -> set_column('K:K',8.5);
$WSOUT_Domi -> set_column('L:L',8.5);
$WSOUT_Domi -> set_column('M:M',4.25);
$WSOUT_Domi -> set_column('N:N',3.38);
$WSOUT_Domi -> set_column('O:O',4.38);
$WSOUT_Domi -> set_column('P:P',4.38);
$WSOUT_Domi -> set_column('Q:Q',4.38);
$WSOUT_Domi -> set_column('R:R',8.38);
$WSOUT_Domi -> set_column('S:S',8.38);
$WSOUT_Domi -> set_column('T:T',8.38);
$WSOUT_Domi -> set_column('U:U',9.13);
$WSOUT_Domi -> set_column('V:V',3.38);
$WSOUT_Domi -> set_column('W:W',3.38);
$WSOUT_Domi -> set_column('X:X',5.88);
$WSOUT_Domi -> set_column('Y:Y',8.38);
$WSOUT_Domi -> set_column('Z:Z',5.88);
$WSOUT_Domi -> set_column('AA:AA',3.38);
foreach $i(0 .. $#ordered_header){
	$WSOUT_Domi->write(0, $i, $ordered_header[$i],$format[$i+1]);
}

my $WSOUT_ComHet = $outworkbook ->add_worksheet(decode("GB2312","ComHet过滤"));
$WSOUT_ComHet -> freeze_panes(0,4);
$WSOUT_ComHet -> set_column('A:A',8.25);
$WSOUT_ComHet -> set_column('B:B',3.25);
$WSOUT_ComHet -> set_column('C:C',11.88);
$WSOUT_ComHet -> set_column('D:D',4.5);
$WSOUT_ComHet -> set_column('E:E',4.25);
$WSOUT_ComHet -> set_column('F:F',9.88);
$WSOUT_ComHet -> set_column('G:G',8.38);
$WSOUT_ComHet -> set_column('H:H',9.0);
$WSOUT_ComHet -> set_column('I:I',9.63);
$WSOUT_ComHet -> set_column('J:J',8.5);
$WSOUT_ComHet -> set_column('K:K',8.5);
$WSOUT_ComHet -> set_column('L:L',8.5);
$WSOUT_ComHet -> set_column('M:M',4.25);
$WSOUT_ComHet -> set_column('N:N',3.38);
$WSOUT_ComHet -> set_column('O:O',4.38);
$WSOUT_ComHet -> set_column('P:P',4.38);
$WSOUT_ComHet -> set_column('Q:Q',4.38);
$WSOUT_ComHet -> set_column('R:R',8.38);
$WSOUT_ComHet -> set_column('S:S',8.38);
$WSOUT_ComHet -> set_column('T:T',8.38);
$WSOUT_ComHet -> set_column('U:U',9.13);
$WSOUT_ComHet -> set_column('V:V',3.38);
$WSOUT_ComHet -> set_column('W:W',3.38);
$WSOUT_ComHet -> set_column('X:X',5.88);
$WSOUT_ComHet -> set_column('Y:Y',8.38);
$WSOUT_ComHet -> set_column('Z:Z',5.88);
$WSOUT_ComHet -> set_column('AA:AA',3.38);
foreach $i(0 .. $#ordered_header){
	$WSOUT_ComHet->write(0, $i, $ordered_header[$i],$format[$i+1]);
}

my $WSOUT_Homo = $outworkbook ->add_worksheet(decode("GB2312","Homo过滤"));
$WSOUT_Homo -> freeze_panes(0,4);
$WSOUT_Homo -> set_column('A:A',8.25);
$WSOUT_Homo -> set_column('B:B',3.25);
$WSOUT_Homo -> set_column('C:C',11.88);
$WSOUT_Homo -> set_column('D:D',4.5);
$WSOUT_Homo -> set_column('E:E',4.25);
$WSOUT_Homo -> set_column('F:F',9.88);
$WSOUT_Homo -> set_column('G:G',8.38);
$WSOUT_Homo -> set_column('H:H',9.0);
$WSOUT_Homo -> set_column('I:I',9.63);
$WSOUT_Homo -> set_column('J:J',8.5);
$WSOUT_Homo -> set_column('K:K',8.5);
$WSOUT_Homo -> set_column('L:L',8.5);
$WSOUT_Homo -> set_column('M:M',4.25);
$WSOUT_Homo -> set_column('N:N',3.38);
$WSOUT_Homo -> set_column('O:O',4.38);
$WSOUT_Homo -> set_column('P:P',4.38);
$WSOUT_Homo -> set_column('Q:Q',4.38);
$WSOUT_Homo -> set_column('R:R',8.38);
$WSOUT_Homo -> set_column('S:S',8.38);
$WSOUT_Homo -> set_column('T:T',8.38);
$WSOUT_Homo -> set_column('U:U',9.13);
$WSOUT_Homo -> set_column('V:V',3.38);
$WSOUT_Homo -> set_column('W:W',3.38);
$WSOUT_Homo -> set_column('X:X',5.88);
$WSOUT_Homo -> set_column('Y:Y',8.38);
$WSOUT_Homo -> set_column('Z:Z',5.88);
$WSOUT_Homo -> set_column('AA:AA',3.38);
foreach $i(0 .. $#ordered_header){
	$WSOUT_Homo->write(0, $i, $ordered_header[$i],$format[$i+1]);
}

my $WSOUT_XL = $outworkbook ->add_worksheet(decode("GB2312","XL过滤"));
$WSOUT_XL -> freeze_panes(0,4);
$WSOUT_XL -> set_column('A:A',8.25);
$WSOUT_XL -> set_column('B:B',3.25);
$WSOUT_XL -> set_column('C:C',11.88);
$WSOUT_XL -> set_column('D:D',4.5);
$WSOUT_XL -> set_column('E:E',4.25);
$WSOUT_XL -> set_column('F:F',9.88);
$WSOUT_XL -> set_column('G:G',8.38);
$WSOUT_XL -> set_column('H:H',9.0);
$WSOUT_XL -> set_column('I:I',9.63);
$WSOUT_XL -> set_column('J:J',8.5);
$WSOUT_XL -> set_column('K:K',8.5);
$WSOUT_XL -> set_column('L:L',8.5);
$WSOUT_XL -> set_column('M:M',4.25);
$WSOUT_XL -> set_column('N:N',3.38);
$WSOUT_XL -> set_column('O:O',4.38);
$WSOUT_XL -> set_column('P:P',4.38);
$WSOUT_XL -> set_column('Q:Q',4.38);
$WSOUT_XL -> set_column('R:R',8.38);
$WSOUT_XL -> set_column('S:S',8.38);
$WSOUT_XL -> set_column('T:T',8.38);
$WSOUT_XL -> set_column('U:U',9.13);
$WSOUT_XL -> set_column('V:V',3.38);
$WSOUT_XL -> set_column('W:W',3.38);
$WSOUT_XL -> set_column('X:X',5.88);
$WSOUT_XL -> set_column('Y:Y',8.38);
$WSOUT_XL -> set_column('Z:Z',5.88);
$WSOUT_XL -> set_column('AA:AA',3.38);
foreach $i(0 .. $#ordered_header){
	$WSOUT_XL->write(0, $i, $ordered_header[$i],$format[$i+1]);
}

%lineNum = ();
$lineNum{Cdd} = 0;
$lineNum{Known} = 0;
$lineNum{ComHet} = 0;
$lineNum{Homo} = 0;
$lineNum{DeNovo} = 0;
$lineNum{Domi} = 0;
$lineNum{XL} = 0;
%count = ();

#open OUT3, $ARGV[2];
open OUT3, $ref_in;
while (<OUT3>){
	next if /^#/;
	chomp;
	@str = split /\t/, $_;
	$cat{$str[0]} = decode("GB2312",$str[2]);
	#$cat{$str[0]} = $str[2];
	$iht{$str[0]} = $str[3];
	if ($str[1] ne ""){
		$cat{$str[1]} = "(".$str[0].")".$str[2];
		$iht{$str[1]} = "(".$str[0].")".$str[3];
	}
}
close OUT3;

%string = ();
%fromF = ();
%fromM = ();

while (<OUT>){
	chomp;
	$all = 0;
	$KG = 0;
	$ExAC = 0;
	$impt = 0;
	@str = split /\t/, $_;
	$gene = $str[6];
	if ($str[9] =~ /(c|n)\.-(\d+)\D/ or $str[9] =~ /(c|n)\.\*(\d+)\D/ or $str[9] =~ /(c|n)\.-(\d+)\_/){
		if ($2 <= 200){
			# print "Yes\n";
			$all = 1;
		}
	}elsif($str[9] =~ /(c|n)\.\d+(\+|\-)(\d+)\D/){
		if ($3 <= 200){
			# print "Yes\n";
			$all = 1;
		}
	}elsif($str[9] =~ /(c|n)\.(\d+)\D/){
		$all = 1;	
		# print "Yes\n";
	}else{
		
	}
	if ($str[11] eq "." or $str[11] <= 0.005){
		$KG = 1;
	}
	
	if ($str[13] eq "." or $str[13] <= 0.005){
		$ExAC = 1;
	}
	next if $KG*$ExAC*$all != 1;
	
	#Candidate
	#push @{$string{Cdd}[$lineNum{Cdd}]}, @str;
	#$lineNum{Cdd} ++;
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		foreach $tmpgene(@tmp){
			if ($str[22] eq "yes"){
				$fromF{$tmpgene} = 1 if $str[21] eq "F";
				$fromM{$tmpgene} = 1 if $str[21] eq "M";
			}
			if ($tmpgene =~ /^(.+)-AS1$/){
				if(exists $cat{$1}){
					$category = $cat{$1};
					$inher = $iht{$1};
				}
				else{
					$category = '';
					$inher = '';
				}
				$count{Cdd}{$1}++;
			}else{
				if(exists $cat{$tmpgene}){
					$category = $cat{$tmpgene};
					$inher = $iht{$tmpgene};
				}
				else{
					$category = '';
					$inher = '';
				}
				$count{Cdd}{$tmpgene}++;
			}
		}
	}elsif($gene =~ /^(.+)-AS1$/){
		if ($str[22] eq "Yes"){
			$fromF{$gene} = 1 if $str[21] eq "F";
			$fromM{$gene} = 1 if $str[21] eq "M";
		}
		if(exists $cat{$1}){
			$category = $cat{$1};
			$inher = $iht{$1};
		}
		else{
			$category = '';
			$inher = '';
		}
		$count{Cdd}{$1}++;
	}else{
		if(exists $cat{$gene}){
			$category = $cat{$gene};
			$inher = $iht{$gene};
		}
		else{
			$category = '';
			$inher = '';
		}
		$count{Cdd}{$gene} ++;
		if ($str[22] eq "Yes"){
			$fromF{$gene} = 1 if $str[21] eq "F";
			$fromM{$gene} = 1 if $str[21] eq "M";
		}
	}
	push @{$string{Cdd}[$lineNum{Cdd}]}, $category, $inher,@str;
	$lineNum{Cdd} ++;
	#Known
	$isKnown = 0;
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		foreach $tmpgene(@tmp){
			if ($tmpgene =~ /^(.+)-AS1$/){
				if(exists $cat{$1} ){
					$category = $cat{$1};
					$inher = $iht{$1};
					$isKnown = 1;
				}
				$count{Known}{$1} ++;
			}
			else{
				if (exists $cat{$tmpgene}){
					$category = $cat{$tmpgene};
					$inher = $iht{$tmpgene};
					$isKnown = 1;
					
				}
				$count{Known}{$tmpgene} ++;
			}
		}
	}else{
		if (exists $cat{$gene}){
			$category = $cat{$gene};
			$inher = $iht{$gene};
			$count{Known}{$gene} ++;
			$isKnown = 1;
		}elsif($gene =~ /^(.+)-AS1$/){
			if(exists $cat{$1} ){
				$category = $cat{$1};
				$inher = $iht{$1};
				
				$isKnown = 1;
			}
			$count{Known}{$1} ++;
		}
	}
	if ($isKnown == 1){
		push @{$string{Known}[$lineNum{Known}]}, $category, $inher, @str;
		$lineNum{Known} ++;
	}
	
	#HOMO
	if ($str[15] eq "'1\/1"){
		if ($str[16] eq "'0\/1" && $str[17] eq "'0\/1"){
			#push @{$string{Homo}[$lineNum{Homo}]}, @str;
			#$lineNum{Homo} ++;
			if ($gene =~ /,/){
				@tmp = split /,/,$gene;
				foreach $tmpgene(@tmp){
					if ($tmpgene =~ /^(.+)-AS1$/){
						if(exists $cat{$1}){
							$category = $cat{$1};
							$inher = $iht{$1};
						}
						else{
							$category = '';
							$inher = '';
						}
						$count{Homo}{$1}++;
					}else{
						if(exists $cat{$tmpgene}){
							$category = $cat{$tmpgene};
							$inher = $iht{$tmpgene};
						}
						else{
							$category = '';
							$inher = '';
						}
						$count{Homo}{$tmpgene}++;
					}
				}
			}elsif($gene =~ /^(.+)-AS1$/){
				if(exists $category{$1}){
					$category = $cat{$1};
					$inher = $iht{$1};
				}
				else{
					$category = '';
					$inher = '';
				}
				$count{Homo}{$1}++;
			}else{
				if(exists $cat{$gene}){
					$category = $cat{$gene};
					$inher = $iht{$gene};
				}
				else{
					$category = '';
					$inher = '';
				}
				$count{Homo}{$gene} ++;
			}
			push @{$string{Homo}[$lineNum{Homo}]}, $category, $inher, @str;
			$lineNum{Homo} ++;
		}
	}
	
	#DENOVO
	if ($str[21] eq "De Novo"){
		$isDenovo = 1;
		($p, $q) = split /\|/,$str[19];
		@tmp = split /:/, $q;
		$isDenovo = 0 if $tmp[1]+ $tmp[2] < 100;
		$isDenovo = 0 if $tmp[2] > 2;
		($p, $q) = split /\|/,$str[20];
		@tmp = split /:/, $q;
		$isDenovo = 0 if $tmp[1]+ $tmp[2] < 100;
		$isDenovo = 0 if $tmp[2] > 2;
		next if $isDenovo == 0;
		#push @{$string{DeNovo}[$lineNum{DeNovo}]}, @str;
		#$lineNum{DeNovo} ++;
		if ($gene =~ /,/){
			@tmp = split /,/,$gene;
			foreach $tmpgene(@tmp){
				if ($tmpgene =~ /^(.+)-AS1$/){
					if(exists $category{$1}){
						$category = $cat{$1};
						$inher = $iht{$1};
					}
					else{
						$category = '';
						$inher = '';
					}
					$count{DeNovo}{$1}++;
				}else{
					if(exists $cat{$tmpgene}){
						$category = $cat{$tmpgene};
						$inher = $iht{$tmpgene};
					}
					else{
						$category = '';
						$inher = '';
					}
					$count{DeNovo}{$tmpgene}++;
				}
			}
		}elsif($gene =~ /^(.+)-AS1$/){
			if(exists $category{$1}){
				$category = $cat{$1};
				$inher = $iht{$1};
			}
			else{
				$category = '';
				$inher = '';
			}
			$count{DeNovo}{$1}++;
		}else{
			if(exists $cat{$gene}){
				$category = $cat{$gene};
				$inher = $iht{$gene};
			}
			else{
				$category = '';
				$inher = '';
			}
			$count{DeNovo}{$gene} ++;
		}
		push @{$string{DeNovo}[$lineNum{DeNovo}]}, $category, $inher, @str;
		$lineNum{DeNovo} ++;
	}
	
	#Dominant
	if($str[15] eq "'0\/1" && $str[16] eq "'0\/0" && $str[17] eq "'0\/0" && $str[0] ne "chrX" && $str[0] ne "chrY"){	#AD
		$isKnown_Domi = 0;
		if ($gene =~ /,/){
			@tmp = split /,/,$gene;
			foreach $tmpgene(@tmp){
				if ($tmpgene =~ /^(.+)-AS1$/){
					if(exists $cat{$1}){
						$category = $cat{$1};
						$inher = $iht{$1};
						$isKnown_Domi = 1;
					}
					$count{Domi}{$1}++;
				}else{
					if(exists $cat{$tmpgene}){
						$category = $cat{$tmpgene};
						$inher = $iht{$tmpgene};
						$isKnown_Domi = 1;
					}
					$count{Domi}{$tmpgene}++;
				}
			}
		}elsif($gene =~ /^(.+)-AS1$/){
			if(exists $category{$1}){
				$category = $cat{$1};
				$inher = $iht{$1};
				$isKnown_Domi = 1;
			}
			$count{Domi}{$1}++;
		}else{
			if(exists $cat{$gene}){
				$category = $cat{$gene};
				$inher = $iht{$gene};
				$isKnown_Domi = 1;
			}
			$count{Domi}{$gene} ++;
		}
		if($isKnown_Domi == 1){
			push @{$string{Domi}[$lineNum{Domi}]}, $category, $inher, @str;
			$lineNum{Domi} ++;
		}
	}
	
	#XL
	if($str[0] eq "chrX" && ( $str[15] eq "'1\/1"|| $str[15] eq "'0\/1" )&& $str[16] eq "'0\/0" && $str[17] eq "'0\/1"){	#XL
		if ($gene =~ /,/){
			@tmp = split /,/,$gene;
			foreach $tmpgene(@tmp){
				if ($tmpgene =~ /^(.+)-AS1$/){
					if(exists $cat{$1}){
						$category = $cat{$1};
						$inher = $iht{$1};
					}
					else{
						$category = '';
						$inher = '';
					}
					$count{Domi}{$1}++;
				}else{
					if(exists $cat{$tmpgene}){
						$category = $cat{$tmpgene};
						$inher = $iht{$tmpgene};
					}
					else{
						$category = '';
						$inher = '';
					}
					$count{XL}{$tmpgene}++;
				}
			}
		}elsif($gene =~ /^(.+)-AS1$/){
			if(exists $category{$1}){
				$category = $cat{$1};
				$inher = $iht{$1};
			}
			else{
				$category = '';
				$inher = '';
			}
			$count{XL}{$1}++;
		}else{
			if(exists $cat{$gene}){
				$category = $cat{$gene};
				$inher = $iht{$gene};
			}
			else{
				$category = '';
				$inher = '';
			}
			$count{XL}{$gene} ++;
		}
		push @{$string{XL}[$lineNum{XL}]}, $category, $inher, @str;
		$lineNum{XL} ++;
	}
}	

close OUT;

# foreach (keys %fromF){
	# print $_,":", $fromF{$_}," | ", $fromM{$_},"\n";
# }

#write Candidate
$i = 0;
foreach $pnt(@{$string{Cdd}}){
	# print $_,"\t" foreach @{$pnt};
	# print "\n";
	$i ++;
	$gene = ${$pnt}[8];
	$isComHet = 0;
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		@tmprepeat = ();
		foreach $tmpgene(@tmp){	
			# print $tmpgene,": ", $fromF{$tmpgene},"|",$fromM{$tmpgene};
			$isComHet = 1 if $fromF{$tmpgene}*$fromM{$tmpgene} == 1;
			if ($tmpgene =~ /^(.+)-AS1$/){
				push @tmprepeat, $count{Cdd}{$1};
			}else{
				push @tmprepeat, $count{Cdd}{$tmpgene};
			}
		}
		unshift @{$pnt}, join (",", @tmprepeat)
	}else{
		# print $gene,": ", $fromF{$gene},"|",$fromM{$gene};
		$isComHet = 1 if $fromF{$gene}*$fromM{$gene} == 1;
		if ($gene =~ /^(.+)-AS1$/){
			unshift @{$pnt}, $count{Cdd}{$1};
		}else{
			unshift @{$pnt}, $count{Cdd}{$gene};
		}
	}
	
	${$pnt}[8] =~ /^([\d\.]+)%/;
	if ($1 <= 20){
		push @{$pnt}, 8;
		$isComHet = 0 if ${$pnt}[-2] ne "Yes";
		push @{$string{ComHet}}, $pnt if $isComHet == 1;
		next;
	}
	if(${$pnt}[12] =~ /(c|n)\.-(\d+)-(\d+)\D/){
		if (($2+$3) > 10){
			push @{$pnt}, 12;
			$isComHet = 0 if ${$pnt}[-2] ne "Yes";
			push @{$string{ComHet}}, $pnt if $isComHet == 1;
			next;
		}
	}elsif (${$pnt}[12] =~ /(c|n)\.-(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.\*(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.-(\d+)\_/){
		if ($2 > 10){
			push @{$pnt}, 12;
			$isComHet = 0 if ${$pnt}[-2] ne "Yes";
			push @{$string{ComHet}}, $pnt if $isComHet == 1;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.\d+(\+|\-)(\d+)\D/){
		if ($3 > 10){
			push @{$pnt}, 12;
			$isComHet = 0 if ${$pnt}[-2] ne "Yes";
			push @{$string{ComHet}}, $pnt if $isComHet == 1;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.(\d+)\D/){
		if (${$pnt}[10] =~ /synonymous|TF_binding/){
			push @{$pnt}, 10;
			$isComHet = 0 if ${$pnt}[-2] ne "Yes";
			push @{$string{ComHet}}, $pnt if $isComHet == 1;
			next;
		}
	}else{
		push @{$pnt}, "no";
		$isComHet = 0 if ${$pnt}[-2] ne "Yes";
		push @{$string{ComHet}}, $pnt if $isComHet == 1;
		next;
	}
	push @{$pnt},"yes";
	
	$isComHet = 0 if ${$pnt}[-2] ne "Yes";
	push @{$string{ComHet}}, $pnt if $isComHet == 1;
}
@tmpyes = ();
@tmpno = ();
foreach $pnt(@{$string{Cdd}}){
	if (${$pnt}[-1] eq "yes"){
		push @tmpyes, $pnt;
	}else{
		push @tmpno, $pnt;
	}
}
@{$string{Cdd}} = ();
push @{$string{Cdd}}, @tmpyes;
push @{$string{Cdd}}, @tmpno;

for(@{$string{Cdd}}){
	@{$_} = &sort_4excel(@{$_});
}
$i = 1;
foreach $pnt(@{$string{Cdd}}){
	# print $_,"\t" foreach @{$pnt};
	# print "\n";
	#$pnt = &sort_pnt(@unsort_header,$pnt);
	&writeline(\$WSOUT_Cdd, $pnt, $i, ${$pnt}[-1]);
	$i ++;
}

#write Known
$i = 0;
foreach $pnt(@{$string{Known}}){
	# print $_,"\t" foreach @{$pnt};
	# print "\n";
	$i ++;
	$gene = ${$pnt}[8];
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		@tmprepeat = ();
		foreach $tmpgene(@tmp){
			if ($tmpgene =~ /^(.+)-AS1$/){
				push @tmprepeat, $count{Known}{$1};
			}else{
				push @tmprepeat, $count{Known}{$tmpgene};
			}
		}
		unshift @{$pnt}, join (",", @tmprepeat)
	}else{
		if ($gene =~ /^(.+)-AS1$/){
			unshift @{$pnt}, $count{Known}{$1};
		}else{
			unshift @{$pnt}, $count{Known}{$gene};
		}
	}
	
	${$pnt}[8] =~ /^([\d\.]+)%/;
	if ($1 <= 20){
		push @{$pnt}, 8;
		next;
	}
	if (${$pnt}[12] =~ /(c|n)\.-(\d+)-(\d+)\D/){
		if (($2+$3) > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif (${$pnt}[12] =~ /(c|n)\.-(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.\*(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.-(\d+)\_/){
		if ($2 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.\d+(\+|\-)(\d+)\D/){
		if ($3 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.(\d+)\D/){
		if (${$pnt}[10] =~ /synonymous|TF_binding/){
			push @{$pnt}, 10;
			next;
		}
	}else{
		push @{$pnt}, "no";
		next;
	}
	push @{$pnt},"yes";
}
@tmpyes = ();
@tmpno = ();
foreach $pnt(@{$string{Known}}){
	if (${$pnt}[-1] eq "yes"){
		push @tmpyes, $pnt;
	}else{
		push @tmpno, $pnt;
	}
}
@{$string{Known}} = ();
push @{$string{Known}}, @tmpyes;
push @{$string{Known}}, @tmpno;

for(@{$string{Known}}){
	@{$_} = &sort_4excel(@{$_});
}
$i = 1;
foreach $pnt(@{$string{Known}}){
	&writeline(\$WSOUT_Known, $pnt, $i, ${$pnt}[-1]);
	$i ++;
}

#write DeNovo
$i = 0;
foreach $pnt(@{$string{DeNovo}}){
	$i ++;
	$gene = ${$pnt}[8];
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		@tmprepeat = ();
		foreach $tmpgene(@tmp){
			if ($tmpgene =~ /^(.+)-AS1$/){
				push @tmprepeat, $count{DeNovo}{$1};
			}else{
				push @tmprepeat, $count{DeNovo}{$tmpgene};
			}
		}
		unshift @{$pnt}, join (",", @tmprepeat)
	}else{
		if ($gene =~ /^(.+)-AS1$/){
			unshift @{$pnt}, $count{DeNovo}{$1};
		}else{
			unshift @{$pnt}, $count{DeNovo}{$gene};
		}
	}
	
	${$pnt}[8] =~ /^([\d\.]+)%/;
	if ($1 <= 20){
		push @{$pnt}, 8;
		next;
	}
	if (${$pnt}[12] =~ /(c|n)\.-(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.\*(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.-(\d+)\_/){
		if ($2 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.\d+(\+|\-)(\d+)\D/){
		if ($3 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.(\d+)\D/){
		if (${$pnt}[10] =~ /synonymous|TF_binding/){
			push @{$pnt}, 10;
			next;
		}
	}else{
		push @{$pnt}, "no";
		next;
	}
	push @{$pnt},"yes";
}
@tmpyes = ();
@tmpno = ();
foreach $pnt(@{$string{DeNovo}}){
	if (${$pnt}[-1] eq "yes"){
		push @tmpyes, $pnt;
	}else{
		push @tmpno, $pnt;
	}
}
@{$string{DeNovo}} = ();
push @{$string{DeNovo}}, @tmpyes;
push @{$string{DeNovo}}, @tmpno;
for(@{$string{DeNovo}}){
	@{$_} = &sort_4excel(@{$_});
}
$i = 1;
foreach $pnt(@{$string{DeNovo}}){
	&writeline(\$WSOUT_Denovo, $pnt, $i, ${$pnt}[-1]);
	$i ++;
}

# write Dominate

$i = 0;
foreach $pnt(@{$string{Domi}}){
	$i ++;
	$gene = ${$pnt}[8];
	#print $_,"\t" foreach @{$pnt};
	#print "\n";
	#next if ${$pnt}[1] !~ /'AD'/;
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		@tmprepeat = ();
		foreach $tmpgene(@tmp){
			if ($tmpgene =~ /^(.+)-AS1$/){
				push @tmprepeat, $count{Domi}{$1};
			}else{
				push @tmprepeat, $count{Domi}{$tmpgene};
			}
		}
		unshift @{$pnt}, join (",", @tmprepeat)
	}else{
		if ($gene =~ /^(.+)-AS1$/){
			unshift @{$pnt}, $count{Domi}{$1};
		}else{
			unshift @{$pnt}, $count{Domi}{$gene};
		}
	}
	
	${$pnt}[8] =~ /^([\d\.]+)%/;
	if ($1 <= 20){
		push @{$pnt}, 8;
		next;
	}
	if (${$pnt}[12] =~ /(c|n)\.-(\d+)-(\d+)\D/){
		if (($2+$3) > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif (${$pnt}[12] =~ /(c|n)\.-(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.\*(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.-(\d+)\_/){
		if ($2 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.\d+(\+|\-)(\d+)\D/){
		if ($3 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.(\d+)\D/){
		if (${$pnt}[10] =~ /synonymous|TF_binding/){
			push @{$pnt}, 10;
			next;
		}
	}else{
		push @{$pnt}, "no";
		next;
	}
	push @{$pnt},"yes";
}
@tmpyes = ();
@tmpno = ();
foreach $pnt(@{$string{Domi}}){
	if (${$pnt}[-1] eq "yes"){
		push @tmpyes, $pnt;
	}else{
		push @tmpno, $pnt;
	}
}
@{$string{Domi}} = ();
push @{$string{Domi}}, @tmpyes;
push @{$string{Domi}}, @tmpno;
for(@{$string{Domi}}){
	@{$_} = &sort_4excel(@{$_});
}
$i = 1;
foreach $pnt(@{$string{Domi}}){
	&writeline(\$WSOUT_Domi, $pnt, $i, ${$pnt}[-1]);
	$i ++;
}

#write ComHet

@tmpyes = ();
@tmpno = ();
%fromF = ();
%fromM = ();
foreach $pnt(@{$string{ComHet}}){
	 #print $_,"\t" foreach @{$pnt};
	 #print "\n";
	if (${$pnt}[-1] eq "yes"){
		$gene = ${$pnt}[9];
		
		if ($gene =~ /,/){
			@tmp = split /,/,$gene;
			foreach $tmpgene(@tmp){
				if (${$pnt}[-2] eq "yes"){
					$fromF{$tmpgene} = 1 if ${$pnt}[-3] eq "F";
					$fromM{$tmpgene} = 1 if ${$pnt}[-3] eq "M";
				}
			}
		}elsif($gene =~ /^(.+)-AS1$/){
			if (${$pnt}[-2] eq "Yes"){
				$fromF{$1} = 1 if ${$pnt}[-3] eq "F";
				$fromM{$1} = 1 if ${$pnt}[-3] eq "M";
			}
		}else{
			if (${$pnt}[-2] eq "Yes"){
				$fromF{$gene} = 1 if ${$pnt}[-3] eq "F";
				$fromM{$gene} = 1 if ${$pnt}[-3] eq "M";
			}
		}
		
		push @tmpyes, $pnt;
	}else{
		${$pnt}[-2] = "no";
		push @tmpno, $pnt;
	}
}
@{$string{ComHet}} = ();
foreach $pnt(@tmpyes){
	$gene = ${$pnt}[9];
	${$pnt}[-2] = "no";
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		foreach $tmpgene(@tmp){
			${$pnt}[-2] = "Yes" if $fromF{$tmpgene}*$fromM{$tmpgene} == 1;
		}
	}elsif($gene =~ /^(.+)-AS1$/){
		${$pnt}[-2] = "Yes" if $fromF{$1}*$fromM{$1} == 1;
	}else{
		${$pnt}[-2] = "Yes" if $fromF{$gene}*$fromM{$gene} == 1;
	}
}
push @{$string{ComHet}}, @tmpyes;
push @{$string{ComHet}}, @tmpno;
#for(@{$string{ComHet}}){
	#print "$_\|" for(@{$_});exit;
#	@{$_} = &sort_4excel(@{$_});
#}
$i = 1;
foreach $pnt(@{$string{ComHet}}){
	&writeline(\$WSOUT_ComHet, $pnt, $i, ${$pnt}[-1]);
	$i ++;
}

#write Homo
$i = 0;
foreach $pnt(@{$string{Homo}}){
	$i ++;
	$gene = ${$pnt}[8];
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		@tmprepeat = ();
		foreach $tmpgene(@tmp){
			if ($tmpgene =~ /^(.+)-AS1$/){
				push @tmprepeat, $count{Homo}{$1};
			}else{
				push @tmprepeat, $count{Homo}{$tmpgene};
			}
		}
		unshift @{$pnt}, join (",", @tmprepeat)
	}else{
		if ($gene =~ /^(.+)-AS1$/){
			unshift @{$pnt}, $count{Homo}{$1};
		}else{
			unshift @{$pnt}, $count{Homo}{$gene};
		}
	}
	
	${$pnt}[8] =~ /^([\d\.]+)%/;
	if ($1 <= 20){
		push @{$pnt}, 8;
		next;
	}
	if (${$pnt}[12] =~ /(c|n)\.-(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.\*(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.-(\d+)\_/){
		if ($2 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.\d+(\+|\-)(\d+)\D/){
		if ($3 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.(\d+)\D/){
		if (${$pnt}[10] =~ /synonymous|TF_binding/){
			push @{$pnt}, 10;
			next;
		}
	}else{
		push @{$pnt}, "no";
		next;
	}
	push @{$pnt},"yes";
}
@tmpyes = ();
@tmpno = ();
foreach $pnt(@{$string{Homo}}){
	if (${$pnt}[-1] eq "yes"){
		push @tmpyes, $pnt;
	}else{
		push @tmpno, $pnt;
	}
}
@{$string{Homo}} = ();
push @{$string{Homo}}, @tmpyes;
push @{$string{Homo}}, @tmpno;
for(@{$string{Homo}}){
	@{$_} = &sort_4excel(@{$_});
}
$i = 1;
foreach $pnt(@{$string{Homo}}){
	&writeline(\$WSOUT_Homo, $pnt, $i, ${$pnt}[-1]);
	$i ++;
}


# write XL

$i = 0;
foreach $pnt(@{$string{XL}}){
	$i ++;
	$gene = ${$pnt}[8];
	#print $_,"\t" foreach @{$pnt};
	#print "\n";
	if ($gene =~ /,/){
		@tmp = split /,/,$gene;
		@tmprepeat = ();
		foreach $tmpgene(@tmp){
			if ($tmpgene =~ /^(.+)-AS1$/){
				push @tmprepeat, $count{XL}{$1};
			}else{
				push @tmprepeat, $count{XL}{$tmpgene};
			}
		}
		unshift @{$pnt}, join (",", @tmprepeat)
	}else{
		if ($gene =~ /^(.+)-AS1$/){
			unshift @{$pnt}, $count{XL}{$1};
		}else{
			unshift @{$pnt}, $count{XL}{$gene};
		}
	}
	
	${$pnt}[8] =~ /^([\d\.]+)%/;
	if ($1 <= 20){
		push @{$pnt}, 8;
		next;
	}
	if (${$pnt}[12] =~ /(c|n)\.-(\d+)-(\d+)\D/){
		if (($2+$3) > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif (${$pnt}[12] =~ /(c|n)\.-(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.\*(\d+)\D/ or ${$pnt}[12] =~ /(c|n)\.-(\d+)\_/){
		if ($2 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.\d+(\+|\-)(\d+)\D/){
		if ($3 > 10){
			push @{$pnt}, 12;
			next;
		}
	}elsif(${$pnt}[12] =~ /(c|n)\.(\d+)\D/){
		if (${$pnt}[10] =~ /synonymous|TF_binding/){
			push @{$pnt}, 10;
			next;
		}
	}else{
		push @{$pnt}, "no";
		next;
	}
	push @{$pnt},"yes";
}
@tmpyes = ();
@tmpno = ();
foreach $pnt(@{$string{XL}}){
	if (${$pnt}[-1] eq "yes"){
		push @tmpyes, $pnt;
	}else{
		push @tmpno, $pnt;
	}
}
@{$string{XL}} = ();
push @{$string{XL}}, @tmpyes;
push @{$string{XL}}, @tmpno;
for(@{$string{XL}}){
	@{$_} = &sort_4excel(@{$_});
}
$i = 1;
foreach $pnt(@{$string{XL}}){
	&writeline(\$WSOUT_XL, $pnt, $i, ${$pnt}[-1]);
	$i ++;
}

$outworkbook -> close();

sub sort_4excel{	#sort_4excel(@str)
	my @str = @_;
	my @order = qw/Gene Repeat GeneCategories Inheritance CHR POS Consequence HGVSc HGVSp KG ESP ExAC Num_of_Harm Origin GT_Case GT_ControlF GT_ControlM Detail_Case Detail_ControlF Detail_ControlM avsnp150 REF ALT FREQ IMPACT ComHet Important/;
	my @sort_tmp =();
	for(@order){
		push @sort_tmp, $str[$index{$_}];
	}
	return @sort_tmp;
}

sub trans_boldnum{	#trans_boldnum($boldNum)
	my @orig_num = qw /9 0 1 2 3 4 10 12 13 14 15 16 17 24 18 19 20 21 22 23 5 6 7 8 11 25/;
	my %trans_hash;
	my $trans_num =0;
	for(@orig_num){
		$trans_hash{$_} = $trans_num;
		$trans_num ++;
	}
	my $bold_num = shift;
	return $trans_hash{$bold_num};
}

sub writeline{
#invoking &writeline($worksheet, \@string, $line, 0/1/other)
#boldNum = "yes": all black
#boldNum = "no": all grey
#boldNum = other: all grey, bold at $boldNum
	$worksheet = shift;
	$substring = shift;
	$line = shift;
	$boldNum = shift;
	
	if ($boldNum eq "yes"){
		foreach $j(0 .. $#{$substring}){
			# print ${$substring}[$j],"\t";
			${$worksheet} -> write($line, $j, ${$substring}[$j],$format2[$j+1]);
		}
		# print "\n";
	}elsif ($boldNum == "no"){
		foreach $j(0 .. $#{$substring}){
			${$worksheet} -> write($line, $j, ${$substring}[$j], $format3[$j+1]);
		}
	}else{
		foreach $j(0 .. (&trans_boldnum($boldNum)-1)){
			${$worksheet} -> write($line, $j, ${$substring}[$j], $format3[$j+1]);
		}
		#print $boldNum,'->',&trans_boldnum($boldNum);exit;
		${$worksheet} -> write($line, &trans_boldnum($boldNum), ${$substring}[&trans_boldnum($boldNum)], $format4[&trans_boldnum($boldNum)+1]);
		foreach $j((&trans_boldnum($boldNum)+1) .. $#{$substring}){
			${$worksheet} -> write($line, $j, ${$substring}[$j], $format3[$j+1]);
		}
	}
}

sub print_usage{
	die "Usage\:\nperl family-screen-candidate-auto_version5.pl [options]
	-I noInterGene.ComHet.family.refined.txt				Required;
	-o Research_sampleID_male/female_trios.xlsx				Required;
	-r epi_1046_genes.info_csh.txt						Required;
	-h help\n";
}