# PhpOfficeTemplate
# [PhpOfficeTemplate](https://www.rightpristine.com/zeikman/PhpOfficeTemplate/demo/) (Beta v0.2.0)

A pure PHP library for filling system data into document/spreadsheet template.
[Demo](https://www.rightpristine.com/zeikman/PhpOfficeTemplate/demo/)

## Get Started

### Installation

~~Via [Composer](https://getcomposer.org/)~~

```shell
# composer require zeikman/phpofficetemplate
```
Installation via Composer is still not working, as I am still figuring out how to publish a composer package, LOL (^o^) !!!

Via [Git](https://github.com/zeikman/PhpOfficeTemplate) *Recommended for the moment*

```shell
# cd <path-to-your-project>
# git clone https://github.com/zeikman/PhpOfficeTemplate
# cd PhpOfficeTemplate
# composer install
```

> **Note!**
>
> You need to install all dependencies manually :
>
> ```shell
> # cd <path-to-your-project>/phpofficetemplate
> # composer install
> ```
>
> If you facing any issue during dependencies Installation, you can try command below :
>
> ```shell
> # composer install --ignore-platform-reqs
> ```

### Dependencies

PhpOfficeTemplate depends on following libraries. Please install all of them using [composer](https://getcomposer.org/).

+ [PhpSpreadsheet 1.25+](https://github.com/PHPOffice/PhpSpreadsheet)
+ [PHPWord 1.0+](https://github.com/PHPOffice/PHPWord)
+ [mPDF 8.1+](https://github.com/mpdf/mpdf/)
+ [TCPDF 6.+6](https://github.com/tecnickcom/TCPDF/)
+ [Dompdf 2.0+](https://github.com/dompdf/dompdf)
+ [Office Converter 1.0](https://github.com/ncjoes/office-converter)

> :warning: **Enable Office Converter!**
> + Office Converter is a PHP Warpper for LibreOffice, in order to use it, you need to install its main dependency, [LibreOffice](http://www.libreoffice.org/).
> + If OfficeConverter does not output any result after install LibreOffice, please try following updates on OfficeConverter lib :
>   1. Go to **vendor/ncjoes/office-converter/src**,
>   2. Open OfficeConverter.php source file using any file editor,
>   3. Change line 245 <br/>**from `$cmd = 'export HOME=/tmp && '.$cmd;`<br/>to `$cmd = 'HOME='.getcwd().' && export HOME && '.$cmd;`**,
>   4. Now try to load your Word/Document template again.

> :information_source: **Tips to install LibreOffice in Linux**
> 1. Download [LibreOffice 7.4.3](https://www.libreoffice.org/download/download-libreoffice/?type=rpm-x86_64&version=7.4.3&lang=en-US) from the official page. <i>(Note: You can download any version that you prefer)</i>
> 2. To [install](https://www.libreoffice.org/get-help/install-howto/linux/) LibreOffice, you are advised to install via the Installation methods recommended by your particular Linux distributon (such as Ubuntu, Centos, and etc). Detailed information is available on the [wiki](https://wiki.documentfoundation.org/Documentation/Install/Linux).

## Usage

### Basic Usage

This would be the simplest way to use PhpOfficeTemplate :
```php
<?php
include_once 'src/PhpOfficeTemplate.php';

// Uploaded file via HTTP POST
$upload_file = $_FILES['upload_file'];

// Configuration
$config = [
  'file_name'  => $upload_file['name'],     // Original name of the uploaded file
  'file_post'  => $upload_file['tmp_name'], // Temporary filename of the uploaded file
  'sheet_name' => 'template'
];

// Creating new template
$template = new PhpOfficeTemplate($config);

// Output template result (default as inline file in PDF)
$template->output();
?>
```

### Configuration

There are two ways of passing the template file into PhpOfficeTemplate :

1. Passing the $_FILES directly into PhpOfficeTemplate using POST method as shown above, and
2. Passing directory path in your server that storing the template file as shown below.
```php
<?php
$config = [
  'file_name'  => $file_name,  // e.g. "template.xlsx"
  'target_dir' => $targer_dir, // e.g. "document/template/"
  'sheet_name' => 'template'
];
?>
```

### Passing Data

To pass the data for variable substitution :
```php
<?php
// Associative array data in [variable => value] format
$data = [
  '${my_variable}'      => 'My Data',
  '${another_variable}' => 'Other Data',
  ...
];

$config = [
  'data' => $data
];
?>
```

### Options

<table>
  <thead>
    <tr>
      <th>Option</th>
      <th>Description</th>
      <th>Default</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>enable_empty_space</td>
      <td>Substitute variable with empty-space if not found.</td>
      <td>false</td>
    </tr>
    <tr>
      <td>enable_office_convertor</td>
      <td>Using OfficeConverter lib for Word/Document output result.</td>
      <td>false</td>
    </tr>
  </tbody>
</table>

### Methods

<table>
  <thead>
    <tr>
      <th>Method</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>getPhpOfficeObject()</td>
      <td>Return PhpOfficeTemplate object.</td>
    </tr>
    <tr>
      <td>setPhpOfficeObject()</td>
      <td>Set/Update PhpOfficeTemplate object.</td>
    </tr>
    <tr>
      <td>setPdfRenderer($pdf_renderer)</td>
      <td>Change PDF renderer.<br/><br/>
        Available options :
        <ul>
          <li>mpdf - Default for Excel/Spreadsheet</li>
          <li>tcpdf - Default for Word/Document</li>
          <li>dompdf</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td>setOrientation($orientation)</td>
      <td>Change page orientation. <i>(ONLY for Excel/Spreadsheet)</i><br/><br/>
        Available options : <i>(Default follow file orientation)</i>
        <ul>
          <li>portrait</li>
          <li>landscape</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td>output([$method, $type, $link])</td>
      <td>Output the result.<br/><br/>
        <i>$method</i> options :<br/>
        <ul>
          <li>browser <i>(Default)</i> - Return a displayable result with following Header :<br/>
            <code>header('Content-type: application/pdf');</code><br/>
            <code>header('Content-Disposition: inline; filename="<i>file_name</i>"');</code><br/>
            <code>header('Cache-Control: max-age=0');</code>
          </li>
          <li>download - Return a downloadable result</li>
          <li>server - Save file to directory in server</li>
        </ul>
        <i>$type</i> options :
        <ul>
          <li>ONLY applicable for <code>'method' => 'download'</code></li>
          <li>Kindly refer to Section <a href="#type-options">$type Options</a></li>
        </ul>
        <i>$link</i> options :
        <ul>
          <li>true - Remove uploaded template after output the result</li>
          <li>false <i>(Default)</i></li>
        </ul>
      </td>
    </tr>
  </tbody>
</table>

### $type Options

ONLY applicable for <code>'method' => 'download'</code>
```php
<?php
...
// Output template result as downloadable PDF file
$template->output([
  'method' => 'download',
  'type' => 'pdf' // default
]);
?>
```

<table>
  <thead>
    <tr>
      <th>Options</th>
      <th>File Type</th>
      <th>Header Returned</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>pdf (Default)</td>
      <td>All file types</td>
      <td>
        <code>header('Content-type: application/pdf');</code><br/>
        <code>header('Content-Disposition: attachment; filename="<i>file_name</i>"');</code><br/>
        <code>header('Cache-Control: max-age=0');</code>
      </td>
    </tr>
    <tr>
      <td>xlsx</td>
      <td>Excel/Spreadsheet</td>
      <td>
        <code>header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');</code><br/>
        <code>header('Content-Disposition: attachment; filename="<i>file_name</i>"');</code><br/>
        <code>header('Cache-Control: max-age=0');</code>
      </td>
    </tr>
    <tr>
      <td>xls</td>
      <td>Excel/Spreadsheet</td>
      <td>
        <code>header('Content-Type: application/vnd.ms-excel');</code><br/>
        <code>header('Content-Disposition: attachment; filename="<i>file_name</i>"');</code><br/>
        <code>header('Cache-Control: max-age=0');</code>
      </td>
    </tr>
    <tr>
      <td>ods</td>
      <td>Excel/Spreadsheet</td>
      <td>
        <code>header('Content-Type: application/vnd.oasis.opendocument.spreadsheet');</code><br/>
        <code>header('Content-Disposition: attachment; filename="<i>file_name</i>"');</code><br/>
        <code>header('Cache-Control: max-age=0');</code>
      </td>
    </tr>
    <tr>
      <td>docx</td>
      <td>Word/Document</td>
      <td>
        <code>header("Content-Description: File Transfer");</code><br/>
        <code>header('Content-Disposition: attachment; filename="<i>file_name</i>"');</code><br/>
        <code>header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');</code><br/>
        <code>header('Content-Transfer-Encoding: binary');</code><br/>
        <code>header('Cache-Control: must-revalidate, post-check=0, pre-check=0');</code><br/>
        <code>header('Expires: 0');</code>
      </td>
    </tr>
    <tr>
      <td>doc</td>
      <td>Word/Document</td>
      <td>
        <code>header("Content-Description: File Transfer");</code><br/>
        <code>header('Content-Disposition: attachment; filename="<i>file_name</i>"');</code><br/>
        <code>header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');</code><br/>
        <code>header('Content-Transfer-Encoding: binary');</code><br/>
        <code>header('Cache-Control: must-revalidate, post-check=0, pre-check=0');</code><br/>
        <code>header('Expires: 0');</code>
      </td>
    </tr>
    <tr>
      <td>odt</td>
      <td>Word/Document</td>
      <td>
        <code>header("Content-Description: File Transfer");</code><br/>
        <code>header('Content-Disposition: attachment; filename="<i>file_name</i>"');</code><br/>
        <code>header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');</code><br/>
        <code>header('Content-Transfer-Encoding: binary');</code><br/>
        <code>header('Cache-Control: must-revalidate, post-check=0, pre-check=0');</code><br/>
        <code>header('Expires: 0');</code>
      </td>
    </tr>
  </tbody>
</table>