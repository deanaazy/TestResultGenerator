#!/usr/bin/perl -w

    use strict;
    use Spreadsheet::ParseExcel;
	use Win32::OLE;
	use Cwd;
	use Win32::OLE::Const 'Microsoft.Word';    # wd  constants
	use Win32::OLE::Const 'Microsoft Office';  # mso constants
	
	my $config_file= "ResultGenerator.cfg";
	
	my $current_path = getcwd;
		$_=$current_path;
	s/\//\\/ig;
	$current_path=$_;
	#print $current_path;
	#my $testcase_file = 'TC_Clinician Portal_Regression Suite_29-Jul-2013_v1.xls';
	#my $testcase_file = 'CMT_RegressionTest_TestCase_V0 4.xls';
	#my $testcase_file = 'CMT_RegressionTest_TestCase_V1 0_Result_20130808.xls';
	my $testcase_file;
	my $testcase_path;
	#$testcase_file = 'TC_Clinician Portal_Regression Suite_29-Jul-2013_v1.xls';
	my $word_app = CreateObject Win32::OLE 'Word.Application' or die $!;
	my $testresult_folder=$current_path.'\\TestResult\\';
	#print $testresult_folder;
	#$word_app->{'Visible'} = 1;

	my $parser = Spreadsheet::ParseExcel->new();
    my $workbook;
	#my $resultFolder = "Result";
	my $testcase_worksheet;
	my $testdata_worksheet;
	my $caseid_col=0;
	my $caseheadline_col=1;
	my $casedes_col=2;
	my $caseexp_col=3;
	my $casedata_col=19;
	my $case_head_row=26;
	my $case_row_start=28;
	my $case_row_end;
	my $data_head_row=0;
	my $data_id_col=0;
	my %data_items;
	my $case_row_min;
	my $case_row_max;
	my $data_row_min;
	my $data_row_max;
	my $casestep;



	
	&main;
	sub main{
	#
	&init;
	&load_conf;
	&open_excel;
	&find_worksheet;
	&validate_format;
	$case_row_min=$case_row_start;
	my ( $case_row_min, $case_row_max ) = $testcase_worksheet->row_range();
	#my $cells;
	print "case row max: $case_row_max\n";
	for my $row ( $case_row_start .. $case_row_max ) {
		#my $cell;
		my $caseid = &get_cell_value($row, $caseid_col);
		#$cell = $testcase_worksheet->get_cell( $row, $caseid_col);
		#if ($cell) {
		#$caseid = $cell->value();
		#}
		my $caseheadline = &get_cell_value( $row, $caseheadline_col );
		my $casedes = &get_cell_value( $row, $casedes_col );
		#$_=$casedes;
		my $caseexp = &get_cell_value( $row, $caseexp_col );		
		my $casedataid = &get_cell_value( $row, $casedata_col );
		$_=$caseid;
		if ($caseid ne ""){
		print "caseid: $caseid\n";
		my $Result_document = $word_app->Documents->Add;
		$Result_document->ActiveWindow->Selection -> TypeText("TestCaseID: $caseid\nHeadLine: $caseheadline\nTestData: $casedataid\n\n");
		my $step;
		my $exp;
		my @casesteps= split /\n/, $casedes;
		my @casesexps= split /\n/, $caseexp;
		my @steps;
		my @exps;
		my $step_ind=-1;
		my $step_flag=0;
		my $exp_ind=-1;
		my $step_id;
		#my $exp_flag=0;		
		
		
		foreach $step (@casesteps){
			#print "Steps: $step\n";
			$_= $step;
			
			if (!/^([a-zA-Z0-9]{1,2})[\.,]/){
				if  ($step_flag==0){
				$Result_document->ActiveWindow->Selection -> TypeText( "$step\n");
				}
			}
			else {
				$step_flag=1;
			}
			if ($step_flag==1){
				if (/^([a-zA-Z0-9]{1,2})[\.,]/){
				$step_ind = $step_ind+1;
				$steps[$step_ind]=$step;
				}
				else{
				$steps[$step_ind]=$steps[$step_ind]."\n".$step;
				}
			}
		}	
			
		foreach $exp (@casesexps){	
			$_=$exp;
			if (/^([a-zA-Z0-9]{1,2})[\.,]/){
				$exp_ind = $exp_ind+1;
				$exps[$exp_ind]=$exp;
			}
			else {
				$exps[$exp_ind]=$exps[$exp_ind]."\n".$exp;
				#print "test";
			}
		}
		
		foreach $step (@steps){
			$_=$step;	
			if (/^([a-zA-Z0-9]{1,2})\./){
				$step_id=$1;
				foreach $exp (@exps){
				#print $exp."\naaaaaaaaaaaaaaaaaaa";
					$_=$exp;
					if (index($exp,($step_id."."))==0){	
						$Result_document->ActiveWindow->Selection -> TypeText( "Step: $step\nExpectedResult: $exp\nScreenShot:\n\n\n");
					}
				}
			}
		}		
		
		$Result_document->SaveAs($testresult_folder."TestResult_$caseid.doc");
		}
		else{
			#print "###################\n";
			#print "CaseID:$caseid.\nHeadline:$caseheadline\nDescription:$casedes\nExpectedResult:$caseexp\nDataID:$casedataid\n###########End#####\n";
			}
			
        }
	$word_app->Quit;	
	}
	# usage: open_excel
	sub open_excel{
		#my $open_excel_parser=shift;
		#my $open_excel_work_book=shift;
		if ( !defined $workbook ) {
			die $parser->error(), ".\n";
		}
	
	}
	
sub load_conf {
	open (PROP,$config_file) || die ("Could not open file: ".$config_file);
	my $line;
	my $value;
	my $prop;
	while ($line=<PROP>){
		$_=$line;
		if (/=/){
			my @item=split /=/, $line;
			$prop=$item[0];
			$value=$item[1];
			chomp ($value);
			if ($prop eq "TestCasePath"){
				$testcase_path=$value;
			}
			if ($prop eq "TestCaseName"){
				$testcase_file=$value;
			}
		}
	}	
	$testcase_file=$testcase_path.$testcase_file;
	if (-e $testcase_file) {print $testcase_file;}
	
	$workbook = $parser->parse($testcase_file);
	close PROP;
}
sub init {
	#clean the Test Result folder
	if (-d $testresult_folder) {
		opendir DIR, ${testresult_folder} or die "Can not open ".$testresult_folder."\n";
		my @old_result_file = readdir DIR;
		my $file;

		foreach $file (@old_result_file) {
			unlink($testresult_folder.$file);
			
		}
		closedir DIR;
	}
	else {mkdir($testresult_folder);}

}	
	
	
	
	sub find_worksheet{
		for my $worksheet ( $workbook->worksheets() ) {
			my $sheetname=$worksheet->get_name();
			$_=$sheetname;
			if (/testcase/i) 
			{
				#print "$sheetname\n";
				$testcase_worksheet=$worksheet;			
			}
			elsif (/test[ ]*data/i)
			{
				#print "$sheetname\n";
				$testdata_worksheet=$worksheet;
			}
			
		}
	
	}
	sub validate_format{

	#my $caseid_col=0;
	#my $caseheadline_col=1;
	#my $casedes_col=2;
	#my $caseexp_col=3;
	#my $casedata_col=19;
	#my $case_head_row=26;
	#my $case_row_start=28;	
	#my $data_head_row=0;
	#my $data_id_col=0;
		my $column_flag=0;
		$_=$testcase_worksheet->get_cell($case_head_row,$caseid_col)->value();
		#print "$_\n";
		if (!/[ ]*/)
		{
			print "caseid head wrong:$_.\n";
			
		}
		$_=$testcase_worksheet->get_cell($case_head_row,$caseheadline_col)->value();
		if (!/Headline/i)
		{
			print "caseheadline head wrong:$_.\n";
			
		}
		$_=$testcase_worksheet->get_cell($case_head_row,$casedes_col)->value();
		if (!/Description/i)
		{
			print "casedes head wrong:$_.\n";
		}
		$_=$testcase_worksheet->get_cell($case_head_row,$caseexp_col)->value();
		if (!/ExpectedResult/i)
		{
			print "caseexp head wrong:$_.\n";
		}
		$_=$testcase_worksheet->get_cell($case_head_row,$casedata_col)->value();
		if (!/TestDataReference/i)
		{
			print "casedata head wrong:$_.\n";
		}
		$_=$testdata_worksheet->get_cell($data_head_row,$data_id_col)->value();
		if (!/Test Data ID/i)
		{
			print "data head wrong:$_.\n";
		}
		$_=$testcase_worksheet->get_cell(($case_row_start-1),$caseid_col)->value();
		if (!/EXAMPLE >>>\nDO NOT DELETE OR USE THIS ROW !!! The 1st row under the column headings row will be deleted during the import process!/i)
		{
			print "test case should not start from row $case_row_start.\n";
			print "$_\n";
			if (/TC/i){
			$case_row_start=$case_row_start-1;
			print "test case should start from row $case_row_start.\n";
			}
		}
	
	}
	# get_cell_value (<row_id>,<column id>);
	sub get_cell_value {
		my $row=shift;
		my $col=shift;
		my $cells;
		my $values;
		$cells = $testcase_worksheet->get_cell( $row, $col);
		if ($cells) {
		$values = $cells->value();
		}
		else {
			$values='';
		}
		return ($values);
	
	}
	
	
	
	
	
	


