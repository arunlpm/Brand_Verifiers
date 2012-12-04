use LWP::UserAgent;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::ParseExcel::Simple;


my $xls = Spreadsheet::ParseExcel::Simple->read('Report.xls');
  foreach my $sheet ($xls->sheets) {

     while ($sheet->has_data) {
         my @data = $sheet->next_row;
         push(@in_brand,$data[0]);

     }
 }

my $sheet=0;
my $parser   = new Spreadsheet::ParseExcel::SaveParser;
my $template = $parser->Parse('Report.xls');
my $format   = $template->{Worksheet}[$sheet]
                            ->{FormatNo};


my $row=1;
for(my $i=1;$i<=$#in_brand;$i++){


	$in_brand[$i]=~s/\s+/\%20/igs;
	$in_brand[$i]="\"".$in_brand[$i]."\"";
	$ua=LWP::UserAgent->new;
	$ua->agent("firefox");
	$url="http://192.168.0.40:8011/solr/select?wt=python&q=title_t:$in_brand[$i]";
	print"$url\n";
	$req=HTTP::Request->new(GET=> $url);
	$res=$ua->request($req);
	$cont=$res->content;
	
	if($cont=~m/\'numFound\':\s*([0-9]*?)\s*\,/is){
	
		my $count=$1;
		print"$count\n";
		$template->AddCell(0, $row,   14, $count,     $format);

	}



$row++;

}

 $template->SaveAs('Report.xls');
