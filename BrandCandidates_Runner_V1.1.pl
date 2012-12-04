use strict;
# use warnings;
# use Data::Dumper;
use MongoDB;
use MongoDB::Collection;
use Spreadsheet::WriteExcel;

my $conn = new MongoDB::Connection( "query_timeout" => "-1");
my $db   = $conn->BrandList;
my $coll = $db->brands;

my $all = $coll->find();


my (@BrandId,@BrandName,@Source,@SOurceCount,@Category,@Categorycount,@InCosmos);
while(my $dts = $all->next)
{

	push(@BrandId,$dts->{BrandId});
	push(@BrandName,$dts->{BrandName});
	push(@Source,$dts->{Source});
	push(@SOurceCount,$dts->{SOurceCount});
	push(@Category,$dts->{Category});
	push(@Categorycount,$dts->{Categorycount});
	push(@InCosmos,$dts->{InCosmos});


}
my $count=$#BrandId;

my $workbook =Spreadsheet::WriteExcel->new('Report.xls');
my $worksheet=$workbook->add_worksheet();
$worksheet->write(0,0,"BrandName");
$worksheet->write(0,1,"Category");
$worksheet->write(0,2,"Source");
$worksheet->write(0,3,"InCosmos");
$worksheet->write(0,4,"Actions");
$worksheet->write(0,5,"SourceCount");
$worksheet->write(0,6,"CategoryCount");
$worksheet->write(0,7,"IdentifiedCategory");
$worksheet->write(0,8,"IdentifiedSource");
$worksheet->write(0,9,"Brand_status");
$worksheet->write(0,10,"Category_status");
$worksheet->write(0,11,"twitter");
$worksheet->write(0,12,"facebook");
$worksheet->write(0,13,"Topics");
$worksheet->write(0,14,"Title_count");
my $r=1;

open(CV,"<test.txt");


while(my $data=<CV>){
	# sleep(2);
	my @arr=split(",",$data);
	my $bName=$arr[0];
	my $category=$arr[2];
	my $source=$arr[1];
	my $incosmos=$arr[3];
	print"***$bName***\n";
		my $catFlg=0;
		for(my $i=0;$i<=$#BrandName;$i++){
			$bName=~s/\s+/\\s\*/igs;
			# print"***$BrandName[$i]***$bName***\n";
			if(defined($BrandName[$i])&& ($BrandName[$i]=~m/^$bName$/is)){

				$bName=~s/\\s\*/ /igs;
				print"******$BrandName[$i]*******\tBrand Matched\n";
				$catFlg=1;

				my @moreCategory=split(";",$category);
				foreach my $subcategory(@moreCategory){
					$subcategory=~s/\s+/\\s\*/igs;
					$subcategory=~s/\)//igs;$subcategory=~s/\(//igs;					
					if($Category[$i]=~m/$subcategory/is){

						$subcategory=~s/\\s\*/ /igs;
						# print"Category Matched\n";
						
						# print".................Matching--$Source[$i]<>$source\n";

						my @moreSource=split(";",$source);
						foreach my $subsource(@moreSource){
							$subsource=~s/\s+/\\s\*/igs;
							if($Source[$i]!~m/$subsource/is){
								$subsource=~s/\\s\*/ /igs;
								my @SrcCount=split(",",$Source[$i]);
								# print scalar(@SrcCount);
								if(scalar(@SrcCount) > 0){
									# print"inside sourcecount\n";
									push(@SrcCount,"$subsource");
									my $srcUp=join(",",@SrcCount);
									my $source_count=scalar(@SrcCount);
									print"UPDATE1\n";
									$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$srcUp","SOurceCount" => "$source_count"}});
									$worksheet->write($r,4,"Source Updated");$worksheet->write($r,5,"$source_count");$worksheet->write($r,8,"$srcUp");
									$Source[$i].=",$subsource";								

								}
								else{
									print"UPDATE2\n";
									$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$subsource"}});
									$worksheet->write($r,4,"Source Updated");
								}
								
							
							}
							else{

								# print"Matched All-Dupe wit Master\n";
								$worksheet->write($r,4,"Duped");
							}
						}
					}

					else{

						my @catCount=split(",",$Category[$i]);
						$subcategory=~s/\\s\*/ /igs;
						if(scalar(@catCount) > 0){

							push(@catCount,"$subcategory");
							my $catup=join(",",@catCount);
							my $category_count=scalar(@catCount);
							print"UPDATE3\n";
							$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Category" => "$catup","CategoryCount" => "$category_count"}});
							$worksheet->write($r,4,"Category Updated");$worksheet->write($r,6,"$category_count");$worksheet->write($r,7,"$catup");
							$Category[$i].=",$subcategory";						
							my @moreSource=split(";",$source);
							foreach my $subsource(@moreSource){
								$subsource=~s/\s+/\\s\*/igs;
								if($Source[$i]!~m/$subsource/is){
									$subsource=~s/\\s\*/ /igs;
									my @SrcCount=split(",",$Source[$i]);
									print scalar(@SrcCount);
									if(scalar(@SrcCount) > 0){
										# print"inside sourcecount\n";
										push(@SrcCount,"$subsource");
										my $srcUp=join(",",@SrcCount);
										my $source_count=scalar(@SrcCount);
										print"UPDATE4\n";
										$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$srcUp","SOurceCount" => "$source_count"}});
										$worksheet->write($r,4,"Source Updated");$worksheet->write($r,5,"$source_count");$worksheet->write($r,8,"$srcUp");
										$Source[$i].=",$subsource";
										
									}
									else{
										# print"Source Not Matched\n";
										print"UPDATE5\n";
										$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$source"}});
										$worksheet->write($r,4,"Source Updated");
										$Source[$i].=",$subsource";
										
									}
								}
							}
						}
						else{
							print"UPDATE6\n";
							$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Category" => "$subcategory"}});
							$worksheet->write($r,4,"Category Updated");
							my @moreSource=split(";",$source);
							foreach my $subsource(@moreSource){
								$subsource=~s/\s+/\\s\*/igs;
								if($Source[$i]!~m/$subsource/is){
									$subsource=~s/\\s\*/ /igs;
									my @SrcCount=split(",",$Source[$i]);
									# print scalar(@SrcCount);
									if(scalar(@SrcCount) > 0){
										# print"inside sourcecount\n";
										push(@SrcCount,"$subsource");
										my $srcUp=join(",",@SrcCount);
										my $source_count=scalar(@SrcCount);
										print"UPDATE7\n";
										$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$srcUp","SOurceCount" => "$source_count"}});
										$worksheet->write($r,4,"Source Updated");$worksheet->write($r,5,"$source_count");$worksheet->write($r,8,"$srcUp");
										$Source[$i].=",$subsource";
										
									}
									else{
										print"UPDATE8\n";
										$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$source"}});
										$worksheet->write($r,4,"Source Updated");
										$Source[$i].=",$subsource";
										
									}
								}
							}
						}
					}
				}
			}

		
		}
		$bName=~s/\\s\*/ /igs;
		$worksheet->write($r,0,"$bName");$worksheet->write($r,1,"$category");$worksheet->write($r,2,"$source");$worksheet->write($r,3,"$incosmos");
		if($catFlg == 0){

			print"Not Found \n inserting a record...\n";
			$count++;
			my $id="BR".$count;
			# print"Last rec-$id\nincremented..\n";
			$id++;
			print"INSERT1\n";
			$db->brands->insert({"BrandId" => "$id","BrandName" => "$bName","Source" => "$source","SOurceCount" => "1","Category" => "$category","CategoryCount" => "1","InCosmos" => "$incosmos"});
			$worksheet->write($r,4,"Added");$worksheet->write($r,5,"1");$worksheet->write($r,6,"1");
			print"Inserted\n";

		}
		$r++;
	

}

 close(CV);