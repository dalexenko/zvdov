<?

// $dbf_file = "C:\BAL_0811.DBF";
// $work_dir = "d:\zvdov";

$dbfile = str_replace("\\", "/" , $dbf_file);
$workdir = str_replace("\\", "/" , $work_dir);

$file_part_tmp = explode("_", $dbfile);
$file_part = explode(".", $file_part_tmp[1]);

$blank_xls_filename = "blank.xls";

$new_xls_filename = $workdir."/xls/BAL_".$file_part[0].".xls";

copy ($blank_xls_filename, $new_xls_filename); 

$date_y = "2010";
$date_m = substr($file_part[0], 0, 2);
$date_d = substr($file_part[0], 2, 4);
$date_str = "на ".$date_d.".".$date_m.".".$date_y; 
$sheet1 = "Лист1";

$excel_app = new COM("Excel.application") or Die ("Did not connect");

$excel_app->Visible = 0;

$Workbook = $excel_app->Workbooks->Open("$new_xls_filename") or Die("Did not open $filename $Workbook");

$i=8;

$coord_date = "F5";

$Worksheet = $Workbook->Worksheets($sheet1);
$Worksheet->activate;

$excel_cell_date = $Worksheet->Range($coord_date);
$excel_cell_date->activate;
$excel_cell_date->value = $date_str;


$excel_result_balacc = '0000';
while ($excel_result_balacc !='')
	{
        $coord_balacc = "B" . $i;
        $coord_DB_DAY = "D" . $i;
        $coord_CR_DAY = "J" . $i;



        $excel_cell_balacc = $Worksheet->Range($coord_balacc);
        $excel_cell_balacc->activate;
        $excel_result_balacc = $excel_cell_balacc->value;

		if($excel_result_balacc != 'Усього')
			{
			
			// open in read-only mode
			$db = dbase_open($dbfile, 0);

			if ($db)
				{
				// read some data ..

				$record_numbers = dbase_numrecords($db);

				for ($y = 1; $y <= $record_numbers; $y++)
					{
					// do something here, for each record

					$row = dbase_get_record_with_names($db, $y);

					if($row['BALANCE'] == $excel_result_balacc)
						{

						$db_day_get_val_command = "cdbflite.exe ".$dbfile." /filter:BALANCE=".$excel_result_balacc." /sum:DB_DAY";
						
						// $db_day_val = system ($db_day_get_val_command);
						
						$cr_day_get_val_command = "cdbflite.exe ".$dbfile." /filter:BALANCE=".$excel_result_balacc." /sum:CR_DAY";
						// $cr_day_val = system ($cr_day_get_val_command);

						$excel_cell_DB_DAY = $Worksheet->Range($coord_DB_DAY);
						$excel_cell_DB_DAY->activate;
						$excel_cell_DB_DAY->value = str_replace(".", "," , substr(trim(exec ($db_day_get_val_command)), 0, -1));
						
						$excel_cell_CR_DAY = $Worksheet->Range($coord_CR_DAY);
						$excel_cell_CR_DAY->activate;
						$excel_cell_CR_DAY->value = str_replace(".", "," , substr(trim (exec ($cr_day_get_val_command)), 0, -1));
						
						}

					}
			dbase_close($db);

				}

			}
	$i = $i + 1;
	}

//$filedest_saved = str_replace("/", "\\" , $new_xls_filename);

$excel_app->ActiveWorkbook->Save();

$excel_app->Quit();

// free the object

//$excel_app->Release();

$excel_app = null;
echo "обработан файл ".$dbf_file;
?>