<?php

namespace GGGGino\WordBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Response;

class FakeController extends Controller
{
    public function streamAction()
    {
        // create an empty object
        $phpExcelObject = $this->createWordObject();
        // create the writer
        $writer = $this->get('phpword')->createWriter($phpExcelObject, 'Word2007');
        // create the response
        $response = $this->get('phpword')->createStreamedResponse($writer);
        // adding headers
        $response->headers->set('Content-Type', 'application/msword');
        $response->headers->set('Content-Disposition', 'attachment;filename=stream-file.doc');
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');

        return $response;
    }

    public function storeAction()
    {
        // create an empty object
        $phpExcelObject = $this->createWordObject();
        // create the writer
        $writer = $this->get('phpword')->createWriter($phpExcelObject, 'Word2007');
        $filename = tempnam(sys_get_temp_dir(), 'doc-') . '.doc';
        // create filename
        $writer->save($filename);

        return new Response($filename, 201);
    }

    public function readerAction()
    {
        $filename = $this->container->getParameter('xls_fixture_absolute_path');
        // load the factory
        /** @var \GGGGino\WordBundle\Factory $reader */
        $factory = $this->get('phpword');
        // create a reader
        /** @var \PHPExcel_Reader_IReader $reader */
        $reader = $factory->createReader('Word2007');
        // check that the file can be read
        $canread = $reader->canRead($filename);
        // check that an empty temporary file cannot be read
        $someFile = tempnam($this->getParameter('kernel.root_dir'), "tmp");
        $cannotread = $reader->canRead($someFile);
        unlink($someFile);
        die(fopen($filename, "r"));
        // load the excel file
        $phpExcelObject = $reader->load($filename);
        // read some data
        $sheet = $phpExcelObject->getActiveSheet();
        $hello = $sheet->getCell('A1')->getValue();
        $world = $sheet->getCell('B2')->getValue();

        return new Response($canread && !$cannotread ? "$hello $world" : 'I should no be able to read this.');
    }

    public function readAndSaveAction()
    {
        $filename = $this->container->getParameter('xls_fixture_absolute_path');
        // create an object from a filename
        $phpExcelObject = $this->createWordObject($filename);
        // create the writer
        $writer = $this->get('phpword')->createWriter($phpExcelObject, 'Word2007');
        $filename = tempnam(sys_get_temp_dir(), 'doc-') . '.doc';
        // create filename
        $writer->save($filename);

        return new Response($filename, 201);
    }

    /**
     * utility class
     * @return mixed
     */
    private function createWordObject()
    {
        $phpWordObject = $this->get('phpword')->createPHPWordObject();

        $section = $phpWordObject->addSection();
        // Adding Text element to the Section having font styled by default...
        $section->addText(
            '"Learn from yesterday, live for today, hope for tomorrow. '
                . 'The important thing is not to stop questioning." '
                . '(Albert Einstein)'
        );

        return $phpWordObject;
    }
}
