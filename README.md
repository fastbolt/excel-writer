# Excel-Writer
This component is used for simple Excel-file generation in Symfony. You can pass either arrays of arrays, or arrays of entities to the generator class. If an entity references another entity, you will need to pass a static function to retrieve a specific value from that entity.

## Table size
The width of the table is set by the number of headers. The number of rows is set by the amount of data.

## Columns
The order of columns is determined by the order of ColumnSetting instances given to ExcelGenerator::setColumns().
You need to define at least one Column.
The order of the content will not be changed if arrays are passed to setContent().

If you are passing objects to setContent(), you will need to provide the name of the method that returns
the values you want to display (like "getName").
```php
    $columns = [
            new ColumnSetting('Name', ColumnSetting::FORMAT_STRING, 'getName'),
            new ColumnSetting('ID', ColumnSetting::FORMAT_INTEGER, 'getId'),
    ];

    $generator = new ExcelGenerator(new SpreadSheetType());
    $file = $generator
                ->setContent($data)
                ->setColumns($columns)
                ->generateSpreadsheet('../var/temp/excelwriter');
```

You can also pass closures instead of the getter.
```php
    $columns = [
            new ColumnSetting('Loginname', ColumnSetting::FORMAT_STRING, static function($user) {
               return $user->getLoginname();
            })
        ];
```

##Style
Create an instance of the TableStyle class and set styles for the header row and the content. Pass the TableStyle to the generator.
```php
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
    
    $style = new TableStyle();
    $style
        ->setHeaderRowHeight(2)
        ->setHeaderStyle($headerStyle)
        ->setDataRowStyle($dataRowStyle);


    $file = $generator
        ->setContent($data)
        ->setColumns($columns)
        ->setStyle($style)
        ->generateSpreadsheet('../var/temp/filename');
```



##Style presets
The following styles are presets, but can be overwritten in the TableStyle class

###header
- borders: medium
- vertical-alignment: center
- horizontal-alignment: center
- color: FF366092 (blue)

## NOTES
* Floats have a preset decimal length of 2 (0.12), but that can be configured with the 4th parameter of the ColumnSetting constructor or its method setDecimalLength().


## Example usage

### Using arrays
```php
    $data = [
        [
            $users[0],            //instance of a user entity
            'Italy',
            new DateTime('NOW')
        ],
        [
            $users[1],           //instance of a user entity
            'France',
            new DateTime('NOW')
        ]
    ];

    //define columns matching the order of the data
    $columns = [
        new ColumnSetting('Login', ColumnSetting::FORMAT_INTEGER, static function($user) {
            return $user->getLoginname();
        }),
        new ColumnSetting('Country', ColumnSetting::FORMAT_STRING),
        new ColumnSetting('Date', ColumnSetting::FORMAT_DATE)
    ];

    //generate
    $generator = new ExcelGenerator(
        new SpreadSheetType()
    );

    $file = $generator
        ->setContent($data)
        ->setColumns($columns)
        ->generateSpreadsheet('../var/temp/filename');
```

### Using objects

```php
    $repo = $this->getDoctrine()->getRepository(User::class);
    $users = $repo->findBy(['client' => 5]);

    //define columns matching the order of the data
    $columns = [
        new ColumnSetting('Login', ColumnSetting::FORMAT_INTEGER, 'getLoginName'),
        new ColumnSetting('Country', ColumnSetting::FORMAT_STRING, static function($user) {
            return $user->getCountry()->getName();
        }),
        new ColumnSetting('Created', ColumnSetting::FORMAT_DATE, 'getCreated')
    ];

    //generate
    $generator = new ExcelGenerator(
        new SpreadSheetType()
    );

    $file = $generator
        ->setContent($users)
        ->setColumns($columns)
        ->generateSpreadsheet('../var/temp/filename');
```

### Full example using objects and adding style
```php
    $repo  = $this->getDoctrine()->getRepository(User::class);
    $users = $repo->findBy(['client' => 5]);

    //define columns matching the order of the data
    $columns = [
        new ColumnSetting('Login', ColumnSetting::FORMAT_INTEGER, 'getLoginName'),
        new ColumnSetting('Country', ColumnSetting::FORMAT_STRING, static function($user) {
            return $user->getCountry()->getName();
        }),
        new ColumnSetting('Created', ColumnSetting::FORMAT_DATE, 'getCreated'),
        new ColumnSetting('Weight', ColumnSetting::FORMAT_FLOAT, 'getWeight', 2)
    ];
    
    //set style
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
    
    $style = new TableStyle();
    $style
        ->setHeaderRowHeight(2)
        ->setHeaderStyle($headerStyle)
        ->setDataRowStyle($dataRowStyle);

    //generate
    $generator = new ExcelGenerator(
        new SpreadSheetType()
    );

    $file = $generator
        ->setContent($users)
        ->setColumns($columns)
        ->setStyle($style)
        ->generateSpreadsheet('../var/temp/filename');

    //download
    $response = new BinaryFileResponse($file->getPath());
    $response->setContentDisposition(ResponseHeaderBag::DISPOSITION_ATTACHMENT);
    return $response;
```
#   e x c e l - w r i t e r
