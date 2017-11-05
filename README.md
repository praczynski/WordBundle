Symfony2 Word bundle
============

This bundle permits you to create, modify and read word objects.

## License

[![License](https://poser.pugx.org/liuggio/ExcelBundle/license.png)](LICENSE)

## Installation

**1**  Add this to your composer.json

``` yaml
    "minimum-stability": "dev",
    "prefer-stable": true,
``` 

**2**  Add to composer.json to the `require` key

``` shell
    $composer require ggggino/wordbundle
``` 

**3** Register the bundle in ``app/AppKernel.php``

``` php
    $bundles = array(
        // ...
        new GGGGino\WordBundle\GGGGinoWordBundle(),
    );
```

## Get started

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
$writer = $this->get('phpword')->createWriter($phpWordObject, 'Word2007');
$writer->save('file.xls');
```

- Create a Word and create a StreamedResponse:

```php
$writer = $this->get('phpword')->createWriter($phpWordObject, 'Word2007');
$response = $this->get('phpword')->createStreamedResponse($writer);
```

## Not Only 'Word2007'

The list of the types are:

1.  'Word2007'
2.  'ODText'
3.  'HTML'
4.  'PDF'
5.  'RTF'

## Example

### Fake Controller

The best place to start is the fake Controller at `Tests/app/Controller/FakeController.php`, that is a working example.

### More example

You could find a lot of examples in the official PHPWord repository https://github.com/PHPOffice/PHPWord/tree/develop/samples

### Create a new doc

``` php
namespace YOURNAME\YOURBUNDLE\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

class DefaultController extends Controller
{

    public function indexAction($name)
    {
        // ask the service for a Word2007
        $phpWordObject = $this->get('phpword')->createPHPWordObject();

        // Create a new Page
        $section = $phpWordObject->addSection();

        // Adding Text element to the Section having font styled by default...
        $section->addText(
            '"Learn from yesterday, live for today, hope for tomorrow. '
                . 'The important thing is not to stop questioning." '
                . '(Albert Einstein)'
        );

        // create the writer
        $writer = $this->get('phpword')->createWriter($phpWordObject, 'Word2007');
        // create the response
        $response = $this->get('phpword')->createStreamedResponse($writer);
        // adding headers
        $dispositionHeader = $response->headers->makeDisposition(
            ResponseHeaderBag::DISPOSITION_ATTACHMENT,
            'stream-file.doc'
        );
        $response->headers->set('Content-Type', 'application/msword');
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        $response->headers->set('Content-Disposition', $dispositionHeader);

        return $response;        
    }
}
```

### Edit a doc

In the template file(docx) variable should be declared as ${var1}, so in the template you can change "var1" value in this way:
```php
$phpTemplateObject->setValue('var1', 'testValue');
```
Complete example

``` php
namespace YOURNAME\YOURBUNDLE\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

class DefaultController extends Controller
{

    public function indexAction($name)
    {
      $fileName = ".../../test.docx";
    
        // ask the service for a Word2007
        $phpTemplateObject = $this->get('phpword')->createTemplateObject($fileName);

        $phpTemplateObject->setValue('test', 'testValue');
        
        $phpWordObject = $this->get('phpword')->getPhpWordObjFromTemplate($phpTemplateObject);

        // create the writer
        $writer = $this->get('phpword')->createWriter($phpWordObject, 'Word2007');
        // create the response
        $response = $this->get('phpword')->createStreamedResponse($writer);
        // adding headers
        $dispositionHeader = $response->headers->makeDisposition(
            ResponseHeaderBag::DISPOSITION_ATTACHMENT,
            'stream-file.docx'
        );
        $response->headers->set('Content-Type', 'application/msword');
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
3. submit a PullRequest