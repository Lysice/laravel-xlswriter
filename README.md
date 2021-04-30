## Laravel-xlswriter 一款基于xlswriter的laravel扩展包
php-xlswriter是一款高性能的excel读写扩展，laravel-xlswriter基于该扩展做了封装，旨在提供一个便于使用的xlswriter的laravel工具包。
[PHP扩展Xlswriter文档](https://xlswriter-docs.viest.me/zh-cn)

#### 如果本扩展帮助到了你 欢迎star。

####如果本扩展有任何问题或有其他想法 欢迎提 issue与pull request。
### XlsWriter扩展介绍
`XlsWriter`是`viest`开发的一款PHP扩展，目前github的 star 数已达到1.6k。开发语言为C语言。以下是官方文档:
```
xlswriter 是一个 PHP C 扩展，可用于在 Excel 2007+ XLSX 文件中读取数据，插入多个工作表，写入文本、数字、公式、日期、图表、图片和超链接。

它具备以下特性：

一、写入
100％兼容的 Excel XLSX 文件
完整的 Excel 格式
合并单元格
定义工作表名称
过滤器
图表
数据验证和下拉列表
工作表 PNG/JPEG 图像
用于写入大文件的内存优化模式
适用于 Linux，FreeBSD，OpenBSD，OS X，Windows
编译为 32 位和 64 位
FreeBSD 许可证
唯一的依赖是 zlib
二、读取
完整读取数据
光标读取数据
按数据类型读取
xlsx 转 CSV
```

### Laravel-xlswriter使用教程
#### 环境要求
- `xlswriter`  1.3.7
-  `PHP` > 7.0
安装请按照`XlsWriter`的官方文档:[安装教程](https://xlswriter-docs.viest.me/zh-cn/an-zhuang)

#### 安装
```
composer require lysice/laravel-xlswriter
```
将`ServiceProvider`加入到app.php中:
```
'providers' => [

        /*
         * Laravel Framework Service Providers...
         */
        ...
        \Lysice\XlsWriter\XlsWriterServiceProvider::class
    ],
```
发布`Facade` 将`Excel`加入到app.php中:
```
    'aliases' => [
            ...
            'Excel' => \Lysice\XlsWriter\Facade\Writer::class
        ],
```
发布`xlswriter.php`配置文件:
```
    php artisan vendor:publish --provider="Lysice\XlsWriter\XlsWriterServiceProvider"
```
#### 配置
本扩展提供如下几个选项:
`extension`: 导出文档的扩展名 目前只支持xlsx与csv
`mode`:Xlswriter的导出方式，可选内存方式与常规方式:
```
 Excel::MODE_NORMAL
 Excel::MODE_MEMORY
```
`export_mode`: 本扩展支持两种导出数据的模式:直接导出与按行导出。
```
Excel::EXPORT_MODE_DATA,
Excel::EXPORT_MODE_ROW,
```


#### 1.命令
##### 1.1 查看xlswriter扩展是否正常安装
```
    php artisan xls:status
```
展示信息如下:
```
laravel-xlsWriter info:
+---------+-----------------------------------+
| version | dev                               |
| author  | lysice<https://github.com/Lysice> |
| docs    |                                   |
+---------+-----------------------------------+
XlsWriter extension status:
+-------------------------------+----------------------------+
| loaded                        | yes                        |
| xlsWriter author              | Jiexing.Wang (wjx@php.net) |
| xlswriter support             | enabled                    |
| Version                       | 1.3.7                      |
| bundled libxlsxwriter version | 1.0.0                      |
| bundled libxlsxio version     | 0.2.27                     |
+-------------------------------+----------------------------+
```
如您的信息展示如上所示，证明您的`cli`环境下本扩展可用。
##### 1.2.创建Exports类
```
    php artisan xls:export XXXXExport array
```
默认在 `App\Exports` 目录下创建出你的`Export`类。
其中 `XXXXExport` 指定你的导出类 后面指定导出类型。
目前支持三种导出类型:
- array
```
    php artisan xls:export XXXXExport array
```
- query
```
    php artisan xls:export XXXXExport query
```

- Collection
```
    php artisan xls:export XXXXExport collection
```

#### 2.导出array
```
<?php

namespace App\Exports;

use Lysice\XlsWriter\Interfaces\FromArray;

class UserExport implements FromArray
{
    /**
     * @return array
     */
    public function array() : array
    {
        return [
            ['哈哈', 'aaa'],
            ['哈哈', 'aaa'],
            ['哈哈', 'aaa'],
            ['哈哈', 'aaa']
        ];
    }

    /**
     * @return array
     */
    public function headers() : array
    {
        return [];
    }
}
```
##### 2.1 设置列标题
如果您想设置每列的列标题 则需要给`headers`方法返回您的`title`数组。
```
/**
     * @return array
     */
    public function headers() : array
    {
        return [];
    }
```

#### 3.直接下载
使用`Excel`的Facade:
```
    Route::get('xls', function() {
        return Excel::download((new \App\Exports\UserExport()));
    });
```
此时访问路由 `xls` 您可以看到下载的文档。
##### 3.1 指定导出文档名称
```
Route::get('xls', function() {
    // 直接下载
    return Excel::download((new \App\Exports\UserExport()), 'test');
});
```
##### 3.2 指定导出类型
Excel::download方法接收第三个参数可以指定导出类型。目前只支持两种类型:xlsx 与csv。
具体见[`xlswriter`官方文档](https://xlswriter-docs.viest.me/)
```
Route::get('xls', function() {
    return Excel::download((new \App\Exports\UserExport()), 'test', \Lysice\XlsWriter\Excel::TYPE_CSV);
});
```
##### 3.3 指定返回的header
Excel::download方法接收第四个参数,可以指定返回header。
```
Route::get('xls', function() {
    return Excel::download((new \App\Exports\UserExport()), 'test', \Lysice\XlsWriter\Excel::TYPE_CSV, ['a' => b]);
});
```
#### 导出到指定目录
```
Excel::store(
        (new \App\Exports\UserExport()),
        'store.xlsx',
        'public',
        \Lysice\XlsWriter\Excel::TYPE_CSV,
    );
```
##### 设置文档选项
```
Excel::store(
        (new \App\Exports\UserExport()),
        'store.xlsx',
        'public',
        \Lysice\XlsWriter\Excel::TYPE_CSV,
        [
            'visibility' => 'private',
        ]
    );
```
##### 4.Traits
Exportable 用在Exports中
```
    return (new \App\Exports\ExportsExport())->download('2021年3月', \Lysice\XlsWriter\Excel::TYPE_CSV, ['x-download' => true]);
```
```
(new \App\Exports\ExportsExport())->store(
        '2021年4月',
        'public',
        \Lysice\XlsWriter\Excel::TYPE_XLSX, 
        ['visibility' => 'private']
    );
```
参数定义在exports中
#### 5.图表
图表支持的类型均定义在 Lysice\XlsWriter\Supports\Chart中以Chart_为前缀的类型
要想在文档中添加图表 需要使得你的export类实现WithCharts契约 然后在charts方法中实现你的配置
```
class ChartsExport implements FromArray, WithCharts{
    
}
```
###### 面积图 直方图 条形图 折线图 圆环图 雷达图的配置类似
如下配置
```
public function charts()
    {
        return [
            [
                'type' => Constants::CHART_AREA_PERCENT,
                'title' => '图表标题',
                'style' => 47,
                'xName' => 'X name',
                'yName' => 'Y name',
                'row' => 0,
                'column' => 3,
                'series' => [
                    [
                        'name' => 'name1',
                        'data' => '=Sheet1!$B$1:$B$6',
                    ],
                    [
                        'name' => 'name2',
                        'data' => '=Sheet1!$A$1:$A$6',
                    ],
                    [
                        'name' => 'name3',
                        'data' => '=Sheet1!$D$1:$D$6',
                    ],
                    [
                        'name' => 'name4',
                        'data' => '=Sheet1!$C$1:$C$6',
                    ]
                ]
            ]
        ];
    }
``` 
##### 饼图
设置的系列数据只有首个元素可应用 因此只需要设置一个数组即可。
##### 
还可以给数据设置分类.指定要取得分类单元格

#### 6.自动过滤
`export`类实现`WithFilter`接口 指定过滤范围的单元格
```
    public function filter(): string
    {
        return 'A1:D1';
    }
```
#### 7.单元格样式
##### 7.1默认单元格样式
文档可以分别为设置默认的单元格样式。
- 1.约定契约
```
 implements WithDefaultFormat
```
- 2.添加样式
```
    /**
     * @return array
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function defaultFormats(): array
    {
        $formatOne = DefaultFormat::create()->border(12)->background(Constants::COLOR_MAGENTA);
        $formatTwo = DefaultFormat::create()->underline( Constants::UNDERLINE_SINGLE)->wrap();
        return [$formatOne, $formatTwo];
    }
```
该方法返回一个DefaultFormat对象数组,当DefaultFormat数量为1时 默认为内容设置样式。若=2 则以数组0作为header样式, 以数组1的对象作为内容样式。
##### 7.2 列样式支持
##### 7.3 单元格样式支持
当前单元格支持的样式分别有:
- 1.bold() 加粗
- 2.italic() 斜体
- 3.border() 边框

Border取值:
```
\Lysice\XlsWriter\Supports\Constants

Constants::BORDER_THIN                // 薄边框风格
Constants::BORDER_MEDIUM              // 中等边框风格
Constants::BORDER_DASHED              // 虚线边框风格
Constants::BORDER_DOTTED              // 虚线边框样式
Constants::BORDER_THICK               // 厚边框风格
Constants::BORDER_DOUBLE              // 双边风格
Constants::BORDER_HAIR                // 头发边框样式
Constants::BORDER_MEDIUM_DASHED       // 中等虚线边框样式
Constants::BORDER_DASH_DOT            // 短划线边框样式
Constants::BORDER_MEDIUM_DASH_DOT     // 中等点划线边框样式
Constants::BORDER_DASH_DOT_DOT        // Dash-dot-dot边框样式
Constants::BORDER_MEDIUM_DASH_DOT_DOT // 中等点划线边框样式
Constants::BORDER_SLANT_DASH_DOT      // 倾斜的点划线边框风格
```
- 4.align(...$align) 对齐
```
Constants::ALIGN_LEFT,
Constants::ALIGN_CENTER,
Constants::ALIGN_RIGHT,
Constants::ALIGN_FILL,
Constants::ALIGN_JUSTIFY,
Constants::ALIGN_CENTER_ACROSS,
Constants::ALIGN_DISTRIBUTED,
Constants::ALIGN_VERTICAL_TOP,
Constants::ALIGN_VERTICAL_BOTTOM,
Constants::ALIGN_VERTICAL_CENTER,
Constants::ALIGN_VERTICAL_JUSTIFY,
Constants::ALIGN_VERTICAL_DISTRIBUTED
分别对应

Format::FORMAT_ALIGN_LEFT;                 // 水平左对齐
Format::FORMAT_ALIGN_CENTER;               // 水平剧中对齐
Format::FORMAT_ALIGN_RIGHT;                // 水平右对齐
Format::FORMAT_ALIGN_FILL;                 // 水平填充对齐
Format::FORMAT_ALIGN_JUSTIFY;              // 水平两端对齐
Format::FORMAT_ALIGN_CENTER_ACROSS;        // 横向中心对齐
Format::FORMAT_ALIGN_DISTRIBUTED;          // 分散对齐
Format::FORMAT_ALIGN_VERTICAL_TOP;         // 顶部垂直对齐
Format::FORMAT_ALIGN_VERTICAL_BOTTOM;      // 底部垂直对齐
Format::FORMAT_ALIGN_VERTICAL_CENTER;      // 垂直剧中对齐
Format::FORMAT_ALIGN_VERTICAL_JUSTIFY;     // 垂直两端对齐
Format::FORMAT_ALIGN_VERTICAL_DISTRIBUTED; // 垂直分散对齐
```
- 5.font($fontName) 字体 参数:字体名称 字体必须存在于本机

- 6.fontColor($fontColor) 设置字体颜色 接收参数:常量或RGB16进制参数
 
可选常量如
```
    文件:\Lysice\XlsWriter\Supports\Constants

    const COLOR_BLACK = \Vtiful\Kernel\Format::COLOR_BLACK;
    const COLOR_BROWN = \Vtiful\Kernel\Format::COLOR_BROWN;
    const COLOR_BLUE = \Vtiful\Kernel\Format::COLOR_BLUE;
    const COLOR_CYAN = \Vtiful\Kernel\Format::COLOR_CYAN;
    const COLOR_GRAY = \Vtiful\Kernel\Format::COLOR_GRAY;
    const COLOR_LIME = \Vtiful\Kernel\Format::COLOR_LIME;
    const COLOR_GREEN = \Vtiful\Kernel\Format::COLOR_GREEN;
    const COLOR_YELLOW = \Vtiful\Kernel\Format::COLOR_YELLOW;
    const COLOR_WHITE = \Vtiful\Kernel\Format::COLOR_WHITE;
    const COLOR_SILVER = \Vtiful\Kernel\Format::COLOR_SILVER;
    const COLOR_RED = \Vtiful\Kernel\Format::COLOR_RED;
    const COLOR_PURPLE = \Vtiful\Kernel\Format::COLOR_PURPLE;
    const COLOR_PINK = \Vtiful\Kernel\Format::COLOR_PINK;
    const COLOR_ORANGE = \Vtiful\Kernel\Format::COLOR_ORANGE;
    const COLOR_NAVY = \Vtiful\Kernel\Format::COLOR_NAVY;
    const COLOR_MAGENTA = \Vtiful\Kernel\Format::COLOR_MAGENTA;
```
- 7.background($backgroundColor, [$pattern]) 设置背景颜色

pattern 可选参数

```
\Lysice\XlsWriter\Supports\Constants
    const PATTERN_NONE = \Vtiful\Kernel\Format::PATTERN_NONE;
    const PATTERN_SOLID = \Vtiful\Kernel\Format::PATTERN_SOLID;
    const PATTERN_MEDIUM_GRAY = \Vtiful\Kernel\Format::PATTERN_MEDIUM_GRAY;
    const PATTERN_DARK_GRAY = \Vtiful\Kernel\Format::PATTERN_DARK_GRAY;
    const PATTERN_LIGHT_GRAY = \Vtiful\Kernel\Format::PATTERN_LIGHT_GRAY;
    const PATTERN_DARK_HORIZONTAL = \Vtiful\Kernel\Format::PATTERN_DARK_HORIZONTAL;
    const PATTERN_DARK_VERTICAL = \Vtiful\Kernel\Format::PATTERN_DARK_VERTICAL;
    const PATTERN_DARK_DOWN = \Vtiful\Kernel\Format::PATTERN_DARK_DOWN;
    const PATTERN_DARK_UP = \Vtiful\Kernel\Format::PATTERN_DARK_UP;
    const PATTERN_DARK_GRID = \Vtiful\Kernel\Format::PATTERN_DARK_GRID;
    const PATTERN_DARK_TRELLIS = \Vtiful\Kernel\Format::PATTERN_DARK_TRELLIS;
    const PATTERN_LIGHT_HORIZONTAL = \Vtiful\Kernel\Format::PATTERN_LIGHT_HORIZONTAL;
    const PATTERN_LIGHT_VERTICAL = \Vtiful\Kernel\Format::PATTERN_LIGHT_VERTICAL;
    const PATTERN_LIGHT_DOWN = \Vtiful\Kernel\Format::PATTERN_LIGHT_DOWN;
    const PATTERN_LIGHT_UP = \Vtiful\Kernel\Format::PATTERN_LIGHT_UP;
    const PATTERN_LIGHT_GRID = \Vtiful\Kernel\Format::PATTERN_LIGHT_GRID;
    const PATTERN_LIGHT_TRELLIS = \Vtiful\Kernel\Format::PATTERN_LIGHT_TRELLIS;
    const PATTERN_GRAY_125 = \Vtiful\Kernel\Format::PATTERN_GRAY_125;
    const PATTERN_GRAY_0625 = \Vtiful\Kernel\Format::PATTERN_GRAY_0625;
``` 
- 8.fontSize($size) 字体大小

- 9.number($format) 设置数字格式

数字格式可选参数
```
    "0.000",
    "#,##0",
    "#,##0.00",
    "0.00"
```
- 10.underline() 设置下划线

- 11.wrap() 设置单元格换行

- 12.strikeout() 设置删除线

#### 8.工作表缩放
- 实现WithZoom接口 注意 返回值需在10-400之间.
```
    public function zoom()
    {
        return 300;
    }
```
#### 9.工作表网格线
- 实现WithGridLine接口
```
    public function gridLine()
    {
        return Constants::GRIDLINES_HIDE_ALL;
    }
```
参数可选:
```
\Lysice\XlsWriter\Supports\Constants
- const GRIDLINES_HIDE_ALL    = 0; // 隐藏打印网格线 和 文档网格线
- const GRIDLINES_SHOW_SCREEN = 1; // 显示网格线
- const GRIDLINES_SHOW_PRINT  = 2; // 显示打印网格线
- const GRIDLINES_SHOW_ALL    = 3; // 显示打印网格线与文档网格线
```
#### 10.行模式
开启xlswriter的行模式,设置`config/xlswriter.php` 中 
```
'export_mode' => Excel::EXPORT_MODE_ROW,
```
则文档会一行一行导出。
##### 10.1
开启行模式后, 可以设置每列单元格的样式
- 1.
```
implements WithColumnFormat
```
```
- 2.实现接口方法 `columnFormats`
/**
     * @return array
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function columnFormats()
    {
        $formatOne = ColumnFormat::create()
            ->setCellType(CellConstants::CELL_TYPE_TEXT)
            ->border(12)
            ->background(\Lysice\XlsWriter\Supports\Constants::COLOR_MAGENTA);
        $formatTwo = ColumnFormat::create()
            ->setOptions(['dateFormat' => "mm/dd/yy"])
            ->setCellType(CellConstants::CELL_TYPE_DATE)
            ->underline( \Lysice\XlsWriter\Supports\Constants::UNDERLINE_SINGLE)
            ->wrap();
        $formatThree = ColumnFormat::create()
            ->setCellType(CellConstants::CELL_TYPE_FORMULA)
            ->italic()
            ->fontSize(12);
        $formatFour = ColumnFormat::create()
            ->setOptions([
                'widthScale' => 1.1,
                'heightScale' => 1.1
            ])
            ->setCellType(CellConstants::CELL_TYPE_IMAGE)
            ->border(1);
        $formatFive = ColumnFormat::create()
            ->setCellType(CellConstants::CELL_TYPE_URL)
            ->border(1);
        return [$formatOne, $formatTwo, $formatThree, $formatFour, $formatFive];
    }
```
`ColumnFormat`方法返回`ColumnFormat`对象的数组。
`ColumnFormat对象` 继承自`DefaultFormat`, 所以可以支持所有格式。
当前支持5种格式,定义在`ColumnFormat` 中:
```
class ColumnFormat extends DefaultFormat {
    public $cellTypes = [
        CellConstants::CELL_TYPE_TEXT, // 文本格式
        CellConstants::CELL_TYPE_DATE, // 日期格式 
        CellConstants::CELL_TYPE_FORMULA, // 公式格式
        CellConstants::CELL_TYPE_IMAGE, // 图片格式
        CellConstants::CELL_TYPE_URL, // url格式
    ];

......
}
```
注意 要设置单元格格式时，若未指定 `CellType` 则会默认按照 `CELL_TYPE_TEXT` 文本格式来处理。
- 1.文本格式 
```
$formatOne = ColumnFormat::create()
            ->setCellType(CellConstants::CELL_TYPE_TEXT)
            ->border(12)
            ->background(\Lysice\XlsWriter\Supports\Constants::COLOR_MAGENTA);
```
- 2.日期格式 可以设置选项 日期格式,通过`setOptions`来实现。如下代码
```
$formatTwo = ColumnFormat::create()
            ->setOptions(['dateFormat' => "mm/dd/yy"])
            ->setCellType(CellConstants::CELL_TYPE_DATE)
            ->underline( \Lysice\XlsWriter\Supports\Constants::UNDERLINE_SINGLE)
            ->wrap();
```
- 3.公式格式
```
$formatThree = ColumnFormat::create()
            ->setCellType(CellConstants::CELL_TYPE_FORMULA)
            ->italic()
            ->fontSize(12);
```
- 4.图片格式, 可以设置缩放比例。通过`setOptions`来实现。如下代码
```
$formatFour = ColumnFormat::create()
            ->setOptions([
                'widthScale' => 1.1,
                'heightScale' => 1.1
            ])
            ->setCellType(CellConstants::CELL_TYPE_IMAGE)
            ->border(1);
```
- 5.URL格式 可以设置文本与提示 ,由于每行的提示文本 `urlTooltip` `urlText`不一样，因此该数据只能在要导入的数据中定义。
如下数据定义:
```
return [
            ['aaa',time(), '=SUM(A2:A3)',public_path('1.jpeg'), 'http://www.baidu.com', 'urlText' => '链接文本1', 'urlTooltip' => '链接提示1'],
            ['aaa',time(), '=SUM(A2:A3)',public_path('1.jpeg'), 'http://www.baidu.com', 'urlText' => '链接文本2', 'urlTooltip' => '链接提示2'],
            ['aaa',time(), '=SUM(A2:A3)',public_path('1.jpeg'), 'http://www.baidu.com', 'urlText' => '链接文本3', 'urlTooltip' => '链接提示3'],
            ['aaa',time(), '=SUM(A2:A3)',public_path('1.jpeg'), 'http://www.baidu.com', 'urlText' => '链接文本4', 'urlTooltip' => '链接提示4'],
            ['aaa',time(), '=SUM(A2:A3)',public_path('1.jpeg'), 'http://www.baidu.com', 'urlText' => '链接文本5', 'urlTooltip' => '链接提示5'],
        ];
```
以上第五列的单元格的文本会被设置成 `urlText`的值。
```
$formatFive = ColumnFormat::create()
            ->setCellType(CellConstants::CELL_TYPE_URL)
            ->border(1);
```
