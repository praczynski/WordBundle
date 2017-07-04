<?php

namespace GGGGino\WordBundle;

use Symfony\Component\HttpFoundation\StreamedResponse;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Writer\WriterInterface;
use PhpOffice\PhpWord\Reader\ReaderInterface;


/**
 * Factory for PhpWord objects, StreamedResponse, and WriterInterface.
 *
 * @package GGGGino\WordBundle
 */
class Factory
{
    private $phpWordIO;

    public function __construct($phpWordIO = '\PhpOffice\PhpWord\IOFactory')
    {
        $this->phpWordIO = $phpWordIO;
    }

    /**
     * Creates an empty PhpWord Object if the filename is empty, otherwise loads the file into the object.
     *
     * @param string $filename
     *
     * @return \PhpWord
     */
    public function createPHPWordObject($filename = null)
    {
        return (null === $filename) ? new PhpWord() : call_user_func(array($this->phpWordIO, 'load'), $filename);
    }

    /**
     * Create a reader
     *
     * @param string $type
     *
     *
     * @return \ReaderInterface
     */
    public function createReader($type = 'Word2007')
    {
        return call_user_func(array($this->phpWordIO, 'createReader'), $type);
    }

    /**
     * Create a writer given the PHPExcelObject and the type,
     *   the type could be one of PHPExcel_IOFactory::$_autoResolveClasses
     *
     * @param \PhpWord $phpWordObject
     * @param string $type
     *
     *
     * @return WriterInterface
     */
    public function createWriter(PhpWord $phpWordObject, $type = 'Word2007')
    {
        return call_user_func(array($this->phpWordIO, 'createWriter'), $phpWordObject, $type);
    }

    /**
     * Stream the file as Response.
     *
     * @param WriterInterface $writer
     * @param int                      $status
     * @param array                    $headers
     *
     * @return StreamedResponse
     */
    public function createStreamedResponse(WriterInterface $writer, $status = 200, $headers = array())
    {
        return new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            },
            $status,
            $headers
        );
    }

    /**
     * Create a PHPExcel Helper HTML Object
     *
     * @return \PHPExcel_Helper_HTML
     * @deprecated
     */
    public function createHelperHTML()
    {
        return new \PHPExcel_Helper_HTML();
    }
}
