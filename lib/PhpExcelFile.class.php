<?php

class PhpExcelFile extends PHPExcel {
  
  const XLS = 'xlsx';
  const CSV = 'csv';
  
  protected $format;
  protected $encoding;
  
  public function __construct($format = null, $encoding = 'UTF-8') { 
    parent::__construct();
    $this->format = $format ? $format : self::XLS;
    $this->encoding = $encoding;
  }
  
  public function generate() {}
  
  public function save($name = null) {
    switch ($this->getFormat()) {
      case self::XLS:
        $writer = new PHPExcel_Writer_Excel2007($this);
        $writer->setOffice2003Compatibility(true);
        break;
      case self::CSV:
        $writer = new PHPExcel_Writer_CSV($this);
        $writer->setDelimiter(';');
        if ('UTF-8' == $this->getEncoding()) { $writer->setUseBOM(true); }
        break;
      default:
        throw new InvalidArgumentException(sprintf('Unknown PhpExcelFile format: %s', $this->getFormat()));
        break;
    }
    
    $writer->save($name ? $name : $this->getName());
  }
  
  public function getName() {
    return $this->getPath() . $this->getFileName();
  }
  
  public function getFileName() {}
  public function getPath() {}
  
  public function getFormat() { return $this->format; }
  public function getEncoding() { return $this->encoding; }
  
  public function getFileExtension() {
    switch ($this->getFormat()) {
      case self::XLS:
        return 'xlsx';
        break;
      case self::CSV:
        return 'csv';
        break;
      default:
        throw new InvalidArgumentException(sprintf('Unknown PhpExcelFile format: %s', $this->getFormat()));
        break;
    }
  }
  
  public function output($filename = null) {
    $this->generate();
    $filename = $filename ? $filename : $this->getName();
    $this->save($filename);
    
    /*
    header('Content-Description: File Transfer');
    if (headers_sent()) {
      die('Some data has already been output to browser, can\'t send CSV/Excel file');
    }
    header('Cache-Control: public, must-revalidate, max-age=0'); // HTTP/1.1
    header('Pragma: public');
    header('Expires: Sat, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
    // force download dialog
    header('Content-Type: application/force-download');
    header('Content-Type: application/octet-stream', false);
    header('Content-Type: application/download', false);
    header(sprintf('Content-Type: %s; encoding=%s', $this->getContentType(), $this->getEncoding()));
    // use the Content-Disposition header to supply a recommended filename
    header('Content-Disposition: attachment; filename="' . basename($this->getName()) . '";');
    header('Content-Transfer-Encoding: binary');
    echo file_get_contents($filename);
    @unlink($filename);
    exit(0);
    */
    

    // Download the report
    $content_type = $this->getContentType();
    header('Content-Description: File Transfer');
    header('Cache-Control: public, must-revalidate, max-age=0'); // HTTP/1.1
    header('Pragma: public');
    header('Expires: Sat, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
    // force download dialog
    header('Content-Type: application/force-download');
    header('Content-Type: application/octet-stream', false);
    header('Content-Type: application/download', false);
    header(sprintf('Content-Type: %s; encoding=%s', $content_type, 'UTF-8'));
    // use the Content-Disposition header to supply a recommended filename
    header('Content-Disposition: attachment; filename="' . basename($filename) . '";');
    header('Content-Transfer-Encoding: binary');
    echo file_get_contents($filename);
    @unlink($filename);
    exit(0);    
  }
  
  protected function getContentType() {    
    switch ($this->getFormat()) {
      case self::XLS:
        // for Excel 2003 use "application/vnd.ms-excel"
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        
      case self::CSV:
        return 'text/csv';
    }
  }
  
  public function setWidths(array $widths) {
    $sheet = $this->getActiveSheet();
    
    foreach ($widths as $column => $width) {
      if ('auto' == $width) {
        $sheet->getColumnDimension($column)->setAutoSize(true);
      }
      else {
        $sheet->getColumnDimension($column)->setWidth($width);  
      }
    }
  }
  
  public function setWidthsByColumn(array $widths) {
    $sheet = $this->getActiveSheet();
    
    foreach ($widths as $column => $width) {
      if ('auto' == $width) {
        $sheet->getColumnDimensionByColumn($column)->setAutoSize(true);
      }
      else {
        $sheet->getColumnDimensionByColumn($column)->setWidth($width);  
      }
    }
  }
}