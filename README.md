Symfony2 Word bundle
============

This bundle permits you to create, modify and read word objects.

## License

[![License](https://poser.pugx.org/liuggio/ExcelBundle/license.png)](LICENSE)


### Version 1.*

If you have installed an old version, and you are happy to use it, you could find documentation and files
in the [tag v1.0.6](https://github.com/liuggio/ExcelBundle/releases/tag/v1.0.6),
[browse the code](https://github.com/liuggio/ExcelBundle/tree/cf0ecbeea411d7c3bdc8abab14c3407afdf530c4).

## Installation

**1**  Add to composer.json to the `require` key

``` shell
    $composer require ggggino/wordbundle
``` 

**2** Register the bundle in ``app/AppKernel.php``

``` php
    $bundles = array(
        // ...
        new GGGGino\WordBundle\GGGGinoWordBundle(),
    );
```

## TL;DR

- Create an empty object:

``` php
$phpWordObject = $this->get('phpword')->createPHPWordObject();
```

- Create an object from a file:

``` php
$phpWordObject = $this->get('phpword')->createPHPWordObject('file.docx');
```

- Create a Word and write to a file given the object:

```php
$writer = $this->get('phpword')->createWriter($phpWordObject, 'Excel5');
$writer->save('file.xls');
```

- Create a Word and create a StreamedResponse:

```php
$writer = $this->get('phpword')->createWriter($phpWordObject, 'Excel5');
$response = $this->get('phpword')->createStreamedResponse($writer);
```

- Create a Excel file with an image:

```php
$writer = $this->get('phpword')->createPHPWordObject();
$writer->setActiveSheetIndex(0);
$activesheet = $writer->getActiveSheet();

$drawingobject = $this->get('phpword')->createPHPExcelWorksheetDrawing();
$drawingobject->setName('Image name');
$drawingobject->setDescription('Image description');
$drawingobject->setPath('/path/to/image');
$drawingobject->setHeight(60);
$drawingobject->setOffsetY(20);
$drawingobject->setCoordinates('A1');
$drawingobject->setWorksheet($activesheet)
```

## Not Only 'Excel5'

The list of the types are:

1.  'Excel5'
2.  'Excel2007'
3.  'Excel2003XML'
4.  'OOCalc'
5.  'SYLK'
6.  'Gnumeric'
7.  'HTML'
8.  'CSV'

## Example

### Fake Controller

The best place to start is the fake Controller at `Tests/app/Controller/FakeController.php`, that is a working example.

### More example

You could find a lot of examples in the official PHPExcel repository https://github.com/PHPOffice/PHPExcel/tree/develop/Examples

### For lazy devs

``` php
namespace YOURNAME\YOURBUNDLE\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

class DefaultController extends Controller
{

    public function indexAction($name)
    {
        // ask the service for a Excel5
       $phpWordObject = $this->get('phpword')->createPHPExcelObject();

       $phpWordObject->getProperties()->setCreator("ggggino")
           ->setLastModifiedBy("David Ginanni")
           ->setTitle("Office Document")
           ->setSubject("Office Document")
           ->setDescription("Test document for Office, generated using PHP classes.")
           ->setKeywords("office openxml php")
           ->setCategory("Test result file");

        // create the writer
        $writer = $this->get('phpword')->createWriter($phpWordObject, 'Excel5');
        // create the response
        $response = $this->get('phpword')->createStreamedResponse($writer);
        // adding headers
        $dispositionHeader = $response->headers->makeDisposition(
            ResponseHeaderBag::DISPOSITION_ATTACHMENT,
            'stream-file.xls'
        );
        $response->headers->set('Content-Type', 'text/vnd.ms-word; charset=utf-8');
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        $response->headers->set('Content-Disposition', $dispositionHeader);

        return $response;        
    }
}
```

## Contribute

1. fork the project
2. clone the repo
3. get the coding standard fixer: `wget http://cs.sensiolabs.org/get/php-cs-fixer.phar`
4. before the PullRequest you should run the coding standard fixer with `php php-cs-fixer.phar fix -v .`