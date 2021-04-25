Laravel-xlswriter 一款基于xlswriter的laravel扩展包
提供主要功能:
1.命令
查看xlswriter扩展是否正常安装
```
    php artisan xls:status
```
#### 导出
##### 1.导出文件
```
$export = [
        ['Rent', 1000],
        ['Gas',  100],
        ['Food', 300],
        ['Gym',  50],
    ];
    return app()->make(\Lysice\XlsWriter\Writer::class)->download($export);
```
##### 2.导出到硬盘某位置
```
app()->make(\Lysice\XlsWriter\Excel::class)->store(new \App\Exports\ExportsExport(), 'index1.xlsx', 'public', \Lysice\XlsWriter\Excel::TYPE_XLSX, [
        'visibility' => 'private',
    ]);
```
##### 3.文件格式
```
return app()->make(\Lysice\XlsWriter\Excel::class)->download(new \App\Exports\ExportsExport(), \Lysice\XlsWriter\Excel::TYPE_CSV, ['cost', 'aaa']);
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
Responsable??
参数定义在exports中
##### 5.图表
图表支持的类型均定义在 Lysice\XlsWriter\Supports\Chart中以Chart_为前缀的类型
要想在文档中添加图表 需要使得你的export类实现WithCharts契约 然后在charts方法中实现你的配置
如 以array数据为例
```
return [
            [2, 40, 30, 79],
            [3, 40, 25, 87],
            [4, 50, 30, 97],
            [5, 30, 10, 20],
            [6, 25, 5, 30],
            [7, 50, 10, 21]
        ];
```
###### 面积图
当前面积图支持三种类型:
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
##### 直方图 条形图 折线图 圆环图 雷达图
##### 饼图
设置的系列数据只有首个元素可应用 因此只需要设置一个数组即可。
##### 
还可以给数据设置分类.指定要取得分类单元格

#### 自动过滤
export类实现`WithFilter`接口 指定过滤范围的单元格即可
```
public function filter(): string
    {
        return 'A1:D1';
    }
```
#### 单元格样式
##### 默认单元格样式
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
当前默认单元格支持的样式分别有:
- bold() 加粗
- italic() 斜体
- border() 边框
```
BORDER_THIN                // 薄边框风格
BORDER_MEDIUM              // 中等边框风格
BORDER_DASHED              // 虚线边框风格
BORDER_DOTTED              // 虚线边框样式
BORDER_THICK               // 厚边框风格
BORDER_DOUBLE              // 双边风格
BORDER_HAIR                // 头发边框样式
BORDER_MEDIUM_DASHED       // 中等虚线边框样式
BORDER_DASH_DOT            // 短划线边框样式
BORDER_MEDIUM_DASH_DOT     // 中等点划线边框样式
BORDER_DASH_DOT_DOT        // Dash-dot-dot边框样式
BORDER_MEDIUM_DASH_DOT_DOT // 中等点划线边框样式
BORDER_SLANT_DASH_DOT      // 倾斜的点划线边框风格
```
- align(...$align) 对齐
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
- font($fontName) 字体 参数:字体名称 字体必须存在于本机
设置字体颜色 接收参数:常量或RGB16进制参数
- fontColor($fontColor) 
可选常量如
```
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
设置背景颜色
- background($backgroundColor, [$pattern])
pattern 可选参数
```
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
- fontSize($size) 字体大小
设置数字格式
- number($format) 
数字格式可选参数
```
    "0.000",
    "#,##0",
    "#,##0.00",
    "0.00"
```
设置下划线
- underline() 
设置单元格换行
- wrap() 
设置删除线
- strikeout() 
#### 工作表缩放
- 实现WithZoom接口 注意 返回值需在10-400之间.
```
    public function zoom()
    {
        return 300;
    }
```
#### 工作表网格线
- 实现WithGridLine接口注意 返回值在0-3之间。
```
    public function gridLine()
    {
        return 0;
    }
```
参数可选:
- const GRIDLINES_HIDE_ALL    = 0; // 隐藏打印网格线 和 文档网格线
- const GRIDLINES_SHOW_SCREEN = 1; // 显示网格线
- const GRIDLINES_SHOW_PRINT  = 2; // 显示打印网格线
- const GRIDLINES_SHOW_ALL    = 3; // 显示打印网格线与文档网格线

##### 5.Query 大数据量的导出
1.xlswriter的内存模式
2.分块导出
3.分事务导出

##### 命令
1.导入
2.导出
3.状态
```
 php artisan xls:status
```
可以看到如下的内容:
```
+-------------------------------+----------------------------+
| loaded                        | yes                        |
| xlsWriter author              | Jiexing.Wang (wjx@php.net) |
| xlswriter support             | enabled                    |
| Version                       | 1.3.7                      |
| bundled libxlsxwriter version | 0.9.4                      |
| bundled libxlsxio version     | 0.2.27                     |
+-------------------------------+----------------------------+
```
当`loaded` 为`yes` 并且`support`为`enabled`时 状态为正常状态。本扩展可用。
