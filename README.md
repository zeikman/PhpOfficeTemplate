# PhpOfficeTemplate
# [PhpOfficeTemplate](https://www.rightpristine.com/zeikman/PhpOfficeTemplate/demo/) (Beta v0.2.0)

<!-- A pure PHP library for filling system data into document template. (Word/Document/Excel/Spreadsheet) -->
PhpOfficeTemplate is a library written in pure PHP that provides the main functionality of filling data into document templates such as Excel and Word, and then outputting the result in PDF, which could be viewed in browser or as a downloadable attachment.

## Useful Case

It is extremely useful in situations like
1. update data content with an ever-changing document template, or
2. a fixed document template with constantly changeable data, or
3. a fixed document template that need to be filled with multiple sets of data.

For case (1), the user can always update their template at any time and upload it to the system to get the final result in PDF format with all data filled in just a few seconds. It aids in completely avoiding the system update request process, such as
- submitting a system ticket to request template updates,
- waiting for the request to be approved,
- waiting for the IT department to update after receiving approval,

And finally got the template updated, but 3-4 days passed. All of these can be done in just a few minutes.

Case (2) will be the most useful situation. When data was changed, the user simply returned to the system and repeated the normal process to obtain the final result in a matter of "clicking." 

Case (3) is the most powerful, as it allows the user to fill multiple sets of data (5, 20, 50, or even 100!) correctly and neatly into template in a matter of minutes, rather than a hand-writing job, which can take several hours.

[Demo](https://www.rightpristine.com/zeikman/PhpOfficeTemplate/demo/)

## Get Started

### Installation

~~Via [Composer](https://getcomposer.org/)~~

```shell
# composer require zeikman/phpofficetemplate
```
~~Installation via Composer is still not working, as I am still figuring out how to publish a composer package, LOL (^o^) !!!~~

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

+ [PhpSpreadsheet 1.25+](https://github.com/PHPOffice/PhpSpreadsheet) - main dependency
+ [PHPWord 1.0+](https://github.com/PHPOffice/PHPWord) - main dependency
+ [mPDF 8.1+](https://github.com/mpdf/mpdf/)
+ [TCPDF 6.6+](https://github.com/tecnickcom/TCPDF/)
+ [Dompdf 2.0+](https://github.com/dompdf/dompdf)
+ [Office Converter 1.0+](https://github.com/ncjoes/office-converter)

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
  'target_dir' => $targer_dir, // e.g. "document/template/" => server directory that storing the template file
  'sheet_name' => 'template'
];
?>
```

### Passing Data

To pass the data for variable substitution :
```php
<?php
// associative array in [variable => value] format
$data = [
  '${var_name_1}' => 'Name 1',
  '${var_name_2}' => 'Name 2',
  '${var_name_3}' => 'Name 3',
  ...
];

// getting from $_FILES
$json = file_get_contents($_FILES['file']['tmp_name']);
$data = json_decode($json, true);

// getting from database
$data = [];
$result = $mysqli->query("SELECT * FROM Table WHERE id = 1");

if ($result->num_rows > 0) {
  $row = $result->fetch_assoc();

  $data['${var_name}'] = $row['name'];
  $data['${var_ages}'] = $row['ages'];
  $data['${var_addr}'] = $row['addr'];
  ...
}

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
    <tr>
      <td>target_dir</td>
      <td>Target directory for template file upload/read.<br/><br/>
        Needed ONLY when you want to upload template to server or using OfficeConverter.
      </td>
      <td>false</td>
    </tr>
    <tr>
      <td>file_post</td>
      <td>Temporary file name of the file in which the uploaded file was stored on server,<br/>
        e.g. $_FILES['upload_file']['tmp_name'].<br/><br/>
        Specified this argument ONLY when you prefer template to be process without being uploaded to server.
      </td>
      <td></td>
    </tr>
    <tr>
      <td>file_name</td>
      <td>Template file name, or the original name of the uploaded file,<br/>
        e.g. $_FILES['upload_file']['name'].
      </td>
      <td></td>
    <tr>
      <td>file_name</td>
      <td>Template file name, or the original name of the uploaded file,<br/>
        e.g. $_FILES['upload_file']['name'].<br/><br/>
        Specified with original name of the uploaded file ONLY when you specified 'file_post' option.
      </td>
      <td></td>
    </tr>
    <tr>
      <td>file_prefix</td>
      <td>File name prefix during output.</td>
      <td>'Template'</td>
    </tr>
    <tr>
      <td>sheet_name</td>
      <td>Excel sheet name that to be convert.</td>
      <td>'template'</td>
    </tr>
    <tr>
      <td>data</td>
      <td>Associative array data which will be used for variable substitution in [variable => value] format.</td>
      <td>[ ]</td>
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
          <li>mpdf - Default for Excel</li>
          <li>tcpdf - Default for Word</li>
          <li>dompdf</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td>setOrientation($orientation)</td>
      <td>Change page orientation. <i>(ONLY for Excel)</i><br/><br/>
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
// Output template result as downloadable PDF attachment
$template->output([
  'method'  => 'download',
  'type'    => 'pdf' // default
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

## Advance Usage

### Merge Files [Excel/WordPDF]

You can enhance the usage to support multiple files (different file types) upload and output as one PDF file (combine into one PDF file). In order to achieve that, you need to include 'PhpOfficeMerger' library and write some file processing coding (convert non-PDF to PDF) before merging. Kindly refer to [Documentation - Usage](https://www.rightpristine.com/zeikman/PhpOfficeTemplate/demo/) for full example.