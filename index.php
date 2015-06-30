<?php

$cliCounter = 1;
$ticketCounter = 1;
$linesCounter = 1;

require_once('./PHPExcel.php');
require_once('./PHPExcel/IOFactory.php');

$excel2 = PHPExcel_IOFactory::createReader('Excel5');
$excel2 = $excel2->load('white.xls');

$excelTickets = PHPExcel_IOFactory::createReader('Excel5');
$excelTickets = $excelTickets->load('white.xls');

$excelLines = PHPExcel_IOFactory::createReader('Excel5');
$excelLines = $excelLines->load('white.xls');


//Declarar conexion mysqli
$mysqli = new mysqli("localhost", "root", "attime7931", "pos");
if ($mysqli->connect_errno) {
    echo "Falló la conexión con MySQL: (" . $mysqli->connect_errno . ") " . $mysqli->connect_error;
}
ob_start();
$resultado = $mysqli->query("SELECT * FROM customers WHERE SEARCHKEY != 'Generico'");
//Abrir el fichero Excel
$resultado->data_seek(0);
while ($fila = $resultado->fetch_assoc()) {
    CreateCliente($fila['SEARCHKEY'], $fila['SEARCHKEY'], $fila['TAXID'], $fila['NAME'], $fila['NAME'], $fila['ADDRESS'], $fila['CITY'], $fila['POSTAL'], $fila['CITY'], $fila['COUNTRY']);
    echo 'Client - '.$fila['SEARCHKEY'];ob_flush();
}
//Write clients
$objWriter = PHPExcel_IOFactory::createWriter($excel2, 'Excel5');
$objWriter->save('PRO.xls');
/*
 * Crea un cliente en el csv de clientes
 */

function CreateCliente($cod, $comptaCode, $nif, $name, $fiscalName, $address, $poblacion, $cp, $provincia, $country) {
    //Set global variables
    global $cliCounter, $excel2;
    //Init
    $excel2->setActiveSheetIndex(0);
    //Write data for the client
    $excel2->getActiveSheet()->setCellValue("A" . $cliCounter, $cod);
    $excel2->getActiveSheet()->setCellValue("B" . $cliCounter, $comptaCode);
    $excel2->getActiveSheet()->setCellValue("C" . $cliCounter, 0);
    $excel2->getActiveSheet()->setCellValue("D" . $cliCounter, $nif);
    $excel2->getActiveSheet()->setCellValue("E" . $cliCounter, $name);
    $excel2->getActiveSheet()->setCellValue("F" . $cliCounter, $name);
    $excel2->getActiveSheet()->setCellValue("G" . $cliCounter, $address);
    $excel2->getActiveSheet()->setCellValue("H" . $cliCounter, $poblacion);
    $excel2->getActiveSheet()->setCellValue("I" . $cliCounter, intval($cp));
    $excel2->getActiveSheet()->setCellValue("J" . $cliCounter, $provincia);
    $cliCounter++;
}

createTicketsForUser();

/*
 * Create tickets
 *
 * @param {String} $cod - The searchKey for the client
 */

function createTicketsForUser() {
    global $excelTickets, $ticketCounter, $mysqli, $excelLines, $linesCounter;
    $resultado = $mysqli->query("SELECT tickets.*, customers.NAME AS CUSTNAME, receipts.*, customers.SEARCHKEY AS SEARCHKEY, SUM(ticketlines.PRICE * ticketlines.UNITS) AS SUBTOTAL FROM `tickets` INNER JOIN customers ON customers.ID = tickets.CUSTOMER INNER JOIN receipts ON receipts.ID = tickets.ID INNER JOIN ticketlines ON ticketlines.TICKET = tickets.id GROUP BY tickets.ID");
    //Abrir el fichero Excel
    $resultado->data_seek(0);
    while ($fila = $resultado->fetch_assoc()) {
        $excelTickets->setActiveSheetIndex(0);
        $excelTickets->getActiveSheet()->setCellValue("A" . $ticketCounter, '1');
        $excelTickets->getActiveSheet()->setCellValue("B" . $ticketCounter, $fila['TICKETID']);
        $excelTickets->getActiveSheet()->setCellValue("C" . $ticketCounter, $fila['TICKETID']);
        $excelTickets->getActiveSheet()->setCellValue("D" . $ticketCounter, $fila['TICKETID']);
        $excelTickets->getActiveSheet()->setCellValue("E" . $ticketCounter, date('d/m/Y', strtotime($fila['DATENEW'])));
        $excelTickets->getActiveSheet()->setCellValue("F" . $ticketCounter, $fila['SEARCHKEY']);
        $excelTickets->getActiveSheet()->setCellValue("G" . $ticketCounter, 2);
        $excelTickets->getActiveSheet()->setCellValue("I" . $ticketCounter, $fila['CUSTNAME']);
        $excelTickets->getActiveSheet()->setCellValue("R" . $ticketCounter, $fila['SUBTOTAL']);
        $ticketCounter++;
        //Get the lines for this invoiceee
        $resultado2 = $mysqli->query("SELECT ticketlines.*, products.NAME FROM `ticketlines` INNER JOIN tickets ON ticketlines.TICKET = tickets.ID INNER JOIN products ON products.ID = ticketlines.PRODUCT WHERE tickets.TICKETID = " . $fila['TICKETID']);
        $resultado2->data_seek(0);
        while ($fila2 = $resultado2->fetch_assoc()) {
            $excelLines->setActiveSheetIndex(0);
            $excelLines->getActiveSheet()->setCellValue("A" . $linesCounter, '1');
            $excelLines->getActiveSheet()->setCellValue("B" . $linesCounter, $fila['TICKETID']);
            $excelLines->getActiveSheet()->setCellValue("C" . $linesCounter, $linesCounter);
            $excelLines->getActiveSheet()->setCellValue("D" . $linesCounter, $fila2['NAME']);
            $excelLines->getActiveSheet()->setCellValue("E" . $linesCounter, $fila2['NAME']);
            $excelLines->getActiveSheet()->setCellValue("F" . $linesCounter, $fila2['UNITS']);
            $excelLines->getActiveSheet()->setCellValue("J" . $linesCounter, $fila2['PRICE']);
            $excelLines->getActiveSheet()->setCellValue("K" . $linesCounter, $fila2['UNITS'] * $fila2['PRICE']);
            $excelLines->getActiveSheet()->setCellValue("N" . $linesCounter, 3);

            $linesCounter++;
        }
    }
    //Write tickets
    $objWriter = PHPExcel_IOFactory::createWriter($excelTickets, 'Excel5');
    $objWriter->save('FRE.xls');
    //Write tickets Lines
    $objWriter = PHPExcel_IOFactory::createWriter($excelLines, 'Excel5');
    $objWriter->save('LFR.xls');
}
