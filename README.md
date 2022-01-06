###changes
- Getters now required
- Callables als Getters zulässig



#-- Excel Writer Package --
This package is used for simple Excel-file generation. You can pass either arrays of arrays, or arrays of entities. If an entity references another entity, getId() will be called on the referenced entity and the return value will be used instead.

composer.json
``` json
"repositories": [
    {
        "type": "vcs",
        "url": "https://devmgmt.fastbolt.com/sleussink/excelwriter"
    }
],
"require": {
    "sleussink/excel-writer": "dev-main",
```


## Table size
The width of the table is set by the number of headers, the number of rows is set by the amount of data.

## Column order
The order of columns depends on the order in which the getters of the object are listed in the class. To be written to
the table, an attribute has to have a getter of the following format: 'get*'.


To set a custom order or select specific columns to be written, you have to set the columns manually in the order you
need and set the getter of the column like so:
```php
    $columns = [
            new ColumnSetting('Name', ColumnSetting::FORMAT_STRING, 'getName'),
            new ColumnSetting('ID', ColumnSetting::FORMAT_INTEGER, 'getId'),
    ];

    $generator = new ExcelGenerator(); //TODO constructor
    $path = $generator
                ->setContent($data)
                ->setColumns($columns)
                ->generateSpreadsheet('../var/temp/excelwriter');
```
Columns that are not specified are not written to the table.

##Style presets
The following styles are presets, but can be overwritten in the TableStyle class

###header
- borders: medium
- vertical-alignment: center
- horizontal-alignment: center

###data rows
- borders: thin

## NOTES
* Referenced entities will be replaced by their ID, provided they have a method called 'getId'.
* Floats have a preset decimal length of 2 (0.12), but that can be configured with the 4th parameter of the ColumnSetting constructor or its method setDecimalLength().


## Example usage

### Simple example using arrays
```php
    $userEntity = new User();

   $data = [
        [
            'branch1',
            'name1',
            new \DateTime('NOW'),
            1,
            $userEntity
        ],
        [
            'branch2',
            'name2',
            new \DateTime('NOW'),
            2,
            $userEntity
        ]
    ];

    //set columns in order of the data
    $columns = [
        new ColumnSetting('Branch', ColumnSetting::FORMAT_STRING),
        new ColumnSetting('Name', ColumnSetting::FORMAT_STRING),
        new ColumnSetting('Start', ColumnSetting::FORMAT_DATE),
        new ColumnSetting('ID', ColumnSetting::FORMAT_INTEGER)
        new ColumnSetting('ID', ColumnSetting::FORMAT_INTEGER //will show the user ID
    ];

    //generate
    $generator = new ExcelGenerator(  //TODO constructor
        new SpreadSheetType(),
        new LetterProvider(),
        new DataConverter()
    );
    $path = $generator
        ->setContent($data)
        ->setColumns($columns)
        ->generateSpreadsheet('../var/temp/excelwriter');
    
    //download
    $response = new BinaryFileResponse($path);
    $response->setContentDisposition(ResponseHeaderBag::DISPOSITION_ATTACHMENT);
    return $response;
```

### Using objects

```php
    $repo = $this->getDoctrine()->getRepository(Feature::class);
    $data = $repo->findAll();
    
    //set columns (required) and customize order by passing getters (optional)
    $columns = [
            new ColumnSetting('Branch', ColumnSetting::FORMAT_STRING, 'getBranch'),
            new ColumnSetting('Name', ColumnSetting::FORMAT_STRING, 'getName'),
            new ColumnSetting('Start', ColumnSetting::FORMAT_DATE, 'getStart'),
            new ColumnSetting('End', ColumnSetting::FORMAT_DATE, 'getEnd'),
            new ColumnSetting('ID', ColumnSetting::FORMAT_INTEGER, 'getId'),
            new ColumnSetting('ID', ColumnSetting::FORMAT_INTEGER, static function(Feature $feature) {
                return $feature->getCreatedBy()->getLastLogin()->format('Y-m-d');            
            })
    ];
    
    //customize styles (optional)
    $headerStyle = [
          'fill' => [
              'fillType' => Fill::FILL_SOLID,
              'color' => array('rgb' => 'FF9933')
          ]
    ];
    $dataRowStyle = [
        'fill' => array(
            'fillType' => Fill::FILL_SOLID,
            'color' => array('rgb' => '66FF66')
        )
    ];
        
    $tableStyle = new TableStyle();
    $tableStyle->setHeaderStyle($headerStyle)
                ->setHeaderRowHeight(2)
                ->setDataRowStyle($dataRowStyle);
    
    //generate
    $generator = new ExcelGenerator( //TODO constructor
        new SpreadSheetType(),
        new LetterProvider(),
        new DataConverter()
    );
    $generator->generateSpreadsheet('/asd.xlsx');
    
    $path = $generator
        ->setColumns($columns)
        ->setContent($data)
        ->setStyle($tableStyle)
        ->generateSpreadsheet('../var/temp/excelwriter');
        
    //download
    $response = new BinaryFileResponse($path);
    $response->setContentDisposition(ResponseHeaderBag::DISPOSITION_ATTACHMENT);
    return $response;
```
#   e x c e l - w r i t e r  
 