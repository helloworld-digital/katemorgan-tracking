<?php
error_reporting(E_ERROR | E_WARNING | E_PARSE | E_NOTICE);
require __DIR__ . '/vendor/autoload.php';
require_once 'excel_reader2.php';

use Automattic\WooCommerce\Client;
use Mailgun\Mailgun;

# Instantiate the client.
$mgClient = new Mailgun('key-87c607234e93f339f8335d317f755bf9');
$domain = "mg.katemorgan.com.au";
$woocommerce_au = new Client(
    'https://katemorgan.com.au', 
    'ck_5ac353158ccf7968ec7d9eaf48872c12a6949aae', 
    'cs_909914ae41c7d2f90c09ba9b09be3bc1d25c4cbc'
);
$woocommerce_nz = new Client(
    'https://katemorgan.co.nz', 
    'ck_2df69b1845646ae709bb5289f26987d15bd783e6', 
    'cs_327b11f3caa848cef774f6dd7b335e24f33d1c37'
);


if($_POST['attachments']){

	$files = json_decode($_POST['attachments']);

	$file = $files[0];

	foreach ($files as $file_key => $file) {
		$excel_file = fopen("files/".$file->name,"w+");
		$log_file = fopen("log/access.log","a");

		fwrite($excel_file,$mgClient->getAttachment($file->url)->http_response_body);
		fclose($excel_file);
		
		$data = new Spreadsheet_Excel_Reader("files/".$file->name);

		fwrite($log_file,date('Y-m-d H:i:s')." ----- Job Start [".$file->name."] ----- \r\n");

		foreach ($data->sheets[0]['cells'] as $row_key => $row) {
			foreach ($row as $cell_key => $cell) {
				
				if($cell == "KM Invoice Number"){
					$start_row = $row_key + 1;
				}
				else if($cell == "Totals:"){
					$end_row = $row_key - 2;
				}
			}
			
		}

		$count = 1;

		$result = [];

		for($i = $start_row; $i <= $end_row; $i++){
			if(!($i%2)){
				$temp_obj = new stdClass();
				$temp_km_invoice_number = str_replace(" (1)","",str_replace("NZ","",$data->sheets[0]['cells'][$i][3]));
				$temp_obj->km_invoice_number = $temp_km_invoice_number;
				$temp_obj->ordered_date = $data->sheets[0]['cells'][$i][4];
				$temp_obj->required_by_date = $data->sheets[0]['cells'][$i][5];
				$temp_obj->consignee = $data->sheets[0]['cells'][$i][6];
				$temp_obj->units_sent = $data->sheets[0]['cells'][$i][7];
				$temp_obj->total_weight = $data->sheets[0]['cells'][$i][8];
				$temp_obj->total_volume = $data->sheets[0]['cells'][$i][10];
				$temp_obj->transport_co = $data->sheets[0]['cells'][$i][12];
			}
			else{
				$temp_obj->address = $data->sheets[0]['cells'][$i][6];
				$temp_obj->transport_ref = $data->sheets[0]['cells'][$i][12];
				array_push($result,$temp_obj);

			}
			$count++;
		}
		
		
		$recipients = [];
		$recipients_object = new stdClass;

		foreach ($result as $order) {

			if($order->transport_co == "STAR TRACK EXPRESS"){
				try {
					$order_details = $woocommerce_au->get('orders/'.$order->km_invoice_number);
				}
				catch(Exception $e){
					$order_details = $woocommerce_nz->get('orders/'.$order->km_invoice_number);
				}
			}
			else{
				try {
					$order_details = $woocommerce_nz->get('orders/'.$order->km_invoice_number);
				}
				catch(Exception $e){
					$order_details = $woocommerce_au->get('orders/'.$order->km_invoice_number);
				}
			}
			$order_details = $order_details["order"];
			
			$print_object = new stdClass;
			
			$print_object->order_id = $order->km_invoice_number;
			$print_object->email = $order_details['customer']['email'];
			$print_object->consignee = $order->consignee;
			$print_object->transport_co = $order->transport_co;
			$print_object->transport_ref = $order->transport_ref;
			$print_object->order_url = $order_details['view_order_url'];
			
			fwrite($log_file,date('Y-m-d H:i:s')." ".$print_object->order_id." ".$print_object->consignee." ".$print_object->email." ".$print_object->transport_co." ".$print_object->transport_ref." ".$print_object->order_url."\r\n");

			array_push($recipients,$print_object->consignee." <".$print_object->email.">");

			$recipients_object->{$print_object->email}->order_id = $print_object->order_id;
			$recipients_object->{$print_object->email}->consignee = $print_object->consignee;
			$recipients_object->{$print_object->email}->transport_co = $print_object->transport_co;
			$recipients_object->{$print_object->email}->transport_ref = $print_object->transport_ref;
			$recipients_object->{$print_object->email}->order_url = $print_object->order_url;

		
			
		}

		// fwrite($log_file, "----- Recipients -----\r\n");
		// fwrite($log_file, implode(",",$recipients)."\r\n");
		// fwrite($log_file, "----- Recipients Variable -----\r\n");
		// fwrite($log_file, print_r($recipients_object,TRUE)."\r\n");

		$result = $mgClient->sendMessage($domain, array(
		   'from'    => 'Kate Morgan <no-reply@katemorgan.com.au>',
		   'to'      => implode(",",$recipients),
		   'subject' => 'Kate Morgan Tracking Details',
		   'text'    => '%recipient.order_id%  %recipient.transport_co% %recipient.transport_ref% %recipient.order_url%',
		   'recipient-variables' => json_encode($recipients_object),
		   'o:testmode' => true
		));

		$result = $mgClient->sendMessage($domain, array(
		    'from'    => 'Kate Morgan <no-reply@katemorgan.com.au>',
		    'to'      => 'Xian Yang Wong <yang@xmarketing.com.au>',
		    'subject' => 'Tracking Details Sent',
		    'text'    => 'Tracking Details Sent Successfully '.date('Y-m-d H:i:s')
		));

		fwrite($log_file,date('Y-m-d H:i:s')." ----- Job End [".$file->name."] ----- \r\n");
		fclose($log_file);	
	}

}


?>