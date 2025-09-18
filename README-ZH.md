# Excel Table Plus

[![](https://img.shields.io/badge/excel__table__plus-0.0.2-blue)](https://pub.dartlang.org/packages/excel_table_plus)
![](https://img.shields.io/badge/Awesome-Flutter-blue)
![](https://img.shields.io/badge/Platform-Android_iOS_Web_Windows_MacOS_Linux-blue)
![](https://img.shields.io/badge/License-MIT-blue)

语言: 简体中文 | [English](README.md)

一个功能增强的 Flutter Excel 风格表格组件，具有单元格合并、调整大小、样式设置等高级功能。该包提供了一个强大且可定制的表格组件，可用于创建类似 Excel 电子表格的复杂数据布局。

可通过自定义单元格以及单元格样式来创建自定义的表格布局，并且可以通过扩展ExcelController来实现自定义功能。

## 预览

| 示例列表 |
|---------|
| ![示例列表](https://raw.githubusercontent.com/bg-1998/excel_table_plus/main/preview/preview1.png) |

| 简易示例 |
|---------|
| ![简易示例](https://raw.githubusercontent.com/bg-1998/excel_table_plus/main/preview/preview_video1.gif) |

| 高级示例 |
|---------|
| ![高级示例](https://raw.githubusercontent.com/bg-1998/excel_table_plus/main/preview/preview_video2.gif) |

| JSON导出/导入示例 |
|------------------|
| ![JSON导出/导入示例](https://raw.githubusercontent.com/bg-1998/excel_table_plus/main/preview/preview_video3.gif) |

| 表格模板样式演示 |
|------------------|
| ![表格模板样式演示](https://raw.githubusercontent.com/bg-1998/excel_table_plus/main/preview/preview2.png) |

## 核心功能

- 创建自定义行列的表格
- 支持单元格合并
- 自定义单元格样式（颜色、字体、对齐方式等）
- 只读单元格支持
- 单元格交互事件回调
- 可配置的单元格尺寸（宽度和高度）
- 可调整行列大小
- 单元格选择管理
- 可滚动的表格视图
- 自定义单元格构建器，完全控制
- 行列标题（序号）
- 可自定义边框和圆角
- 支持自定义单元格模型类和JSON序列化
- 支持导出和导入表格数据为JSON格式

## 可实现的高级功能

基于核心功能，开发者可以实现以下高级功能：
- 多单元格选择（支持拖拽）
- 插入/删除行列
- 单元格格式化（文本、数字、日期、下拉框）
- 缩放功能
- 撤销/重做操作
- 复杂单元格类型（下拉框、日期选择器等）

## 安装

### 添加到你的项目中

```yaml
dependencies:
  flutter:
    sdk: flutter

  excel_table_plus: ^latest
```

### 安装到项目中

```shell
$ flutter pub get
```

### 导入

```dart
import 'package:excel_table_plus/excel_table_plus.dart';
```

## 使用方法

### 基本用法

```dart
ExcelController controller = ExcelController(
  excel: ExcelModel(
    x: 10,
    y: 20,
    showSn: true,
    backgroundColor: Colors.white,
    rowColor: Colors.blue.withOpacity(.25),
    resizable: true,
    borderRadius: 8.0,
  ),
  items: [
    // 标题行
    ExcelItemModel(
      position: ExcelPosition(0, 0),
      value: '姓名',
      color: Colors.grey.shade200,
      isReadOnly: true,
    ),
    ExcelItemModel(
      position: ExcelPosition(1, 0),
      value: '年龄',
      color: Colors.grey.shade200,
      isReadOnly: true,
    ),
    // 数据单元格
    ExcelItemModel(
      position: ExcelPosition(0, 1),
      value: '张三',
    ),
    ExcelItemModel(
      position: ExcelPosition(1, 1),
      value: '30',
    ),
    // 合并单元格
    ExcelItemModel(
      position: ExcelPosition(0, 5),
      value: '合并单元格示例',
      color: Colors.green,
      isReadOnly: true,
      isMergeCell: true,
      positions: [
        ExcelPosition(0, 5),
        ExcelPosition(1, 5),
        ExcelPosition(2, 5),
        ExcelPosition(0, 6),
        ExcelPosition(1, 6),
        ExcelPosition(2, 6),
      ],
    ),
  ],
);

FlutterExcelWidget(
  controller: controller,
  itemBuilder: (x, y, item) {
    // 自定义单元格部件构建器
    if (item != null) {
      return TextField(
        controller: TextEditingController()..text = item.value?.toString() ?? '',
        readOnly: item.isReadOnly,
        decoration: const InputDecoration(
          border: InputBorder.none,
          contentPadding: EdgeInsets.all(8),
        ),
        onChanged: (value) {
          item.value = value;
        },
      );
    }
    return const SizedBox();
  },
)
```

### 支持JSON序列化的自定义单元格模型

库支持带有JSON序列化的自定义单元格模型类。为了正确地将单元格数据导出和导入为JSON，自定义单元格值类必须实现 `toJson()` 和 `fromJson()` 方法：

```dart
// 带JSON支持的自定义单元格值类
class CustomCellValue {
  String? value;
  final String cellType = 'custom_text';
  final TextEditingController controller;
  final TextAlign textAlign;
  final TextStyle? style;
  
  CustomCellValue({
    this.value,
    TextEditingController? controller,
    this.textAlign = TextAlign.start,
    this.style,
  }) : controller = controller ?? TextEditingController();
  
  // JSON序列化
  Map<String, dynamic> toJson() {
    final Map<String, dynamic> json = <String, dynamic>{};
    json['value'] = value;
    json['cellType'] = cellType;
    json['textAlign'] = textAlign.toString();
    // 如需要，序列化样式
    if (style != null) {
      json['style'] = {
        'fontSize': style!.fontSize,
        'color': style!.color?.value,
        // 根据需要添加其他样式属性
      };
    }
    return json;
  }
  
  // JSON反序列化
  factory CustomCellValue.fromJson(Map<String, dynamic> json) {
    TextAlign textAlign = TextAlign.start;
    // 从字符串解析textAlign
    String textAlignString = json['textAlign'] as String;
    // 根据需要添加解析逻辑
    
    TextStyle? style;
    // 如有可用，解析样式
    if (json['style'] != null) {
      Map<String, dynamic> styleJson = json['style'] as Map<String, dynamic>;
      style = TextStyle(
        fontSize: styleJson['fontSize'] as double?,
        color: styleJson['color'] != null ? Color(styleJson['color'] as int) : null,
        // 根据需要添加其他样式属性
      );
    }
    
    return CustomCellValue(
      value: json['value'] as String?,
      textAlign: textAlign,
      style: style,
    );
  }
}

// 在ExcelItemModel中使用自定义单元格模型
ExcelItemModel(
  position: ExcelPosition(0, 0),
  value: CustomCellValue(value: '自定义单元格值'),
  cellType: 'custom_text', // 指定单元格类型以进行正确的序列化
)
```

### 导出和导入表格数据

库支持将整个表格数据导出为JSON格式并重新导入。导入时，需要提供一个工厂函数来根据类型创建自定义单元格值：

```dart
// 将表格数据导出为JSON
Map<String, dynamic> exportTableData() {
  return {
    'excel': controller.excel.toJson(),
    'items': controller.items.map((item) => item.toJson()).toList(),
  };
}

// 从JSON导入表格数据
void importTableData(Map<String, dynamic> json) {
  ExcelModel excel = ExcelModel.fromJson(json['excel'] as Map<String, dynamic>);
  List<ExcelItemModel> items = (json['items'] as List<dynamic>)
      .map((itemJson) => ExcelItemModel.fromJson(
            itemJson as Map<String, dynamic>,
            customValueFactory: (cellType, valueJson) {
              // 根据类型创建自定义单元格值
              switch (cellType) {
                case 'custom_text':
                  return CustomCellValue.fromJson(valueJson);
                // 为其他自定义单元格类型添加更多情况
                default:
                  return valueJson;
              }
            },
          ))
      .toList();
  
  controller = ExcelController(
    excel: excel,
    items: items,
  );
}
```

### 核心组件

#### ExcelModel

表格的主要配置：

- `x` - 列数
- `y` - 行数
- `itemWidth` - 默认列宽
- `itemHeight` - 默认行高
- `customColumnWidths` - 特定列的自定义宽度
- `customRowHeights` - 特定行的自定义高度
- `selectedBorderWidth` - 选中单元格边框宽度
- `selectedBorderColor` - 选中单元格边框颜色
- `dividerWidth` - 分割线宽度
- `borderWidth` - 表格边线宽度
- `borderRadius` - 表格边线圆角
- `borderColor` - 表格边线颜色
- `rowColor` - 行背景色
- `columnColor` - 列背景色
- `backgroundColor` - 表格背景色
- `dividerColor` - 分割线颜色
- `intersectionColor` - 横竖交集的颜色
- `isReadOnly` - 是否只读，此处为只读的时候，单元格的 isReadOnly 属性将无效
- `isEnableMultipleSelection` - 是否启用多选
- `showSn` - 显示/隐藏序号
- `sn` - 序号配置
- `alignment` - 对齐方式
- `resizable` - 启用/禁用调整大小功能
- `resizeAreaSize` - 调整区域大小（像素）

#### ExcelItemModel

表示表格中的单元格：

- `position` - 单元格位置 (x, y)
- `value` - 单元格值（可以是任何类型）
- `isMergeCell` - 是否为合并单元格
- `positions` - 合并单元格中的所有位置
- `cellType` - 自定义单元格模型类型的字符串标识符
- `color` - 单元格背景色
- `isReadOnly` - 单元格是否只读

#### ExcelController

用于管理表格状态和操作的控制器：

- 提供表格操作方法
- 管理单元格选择状态
- 处理单元格合并操作
- 支持事件回调注册
- 支持查找指定位置的项
- 支持行列的插入和删除
- 支持多选和框选功能
- 支持行列宽高调整

## 示例项目

包中包含了全面的示例，演示了各种实现方式：

1. **简单示例** - 基本用法和简单数据
2. **高级示例** - 复杂用法，包含不同类型的单元格（文本、数字、日期、下拉框）
3. **JSON导出/导入示例** - 演示如何导出和导入带有自定义单元格模型的表格数据
4. **模板示例** - 各种表格样式模板示例

运行示例项目：

```shell
cd example
flutter run
```

## 贡献

欢迎贡献！请随时提交问题和拉取请求。

## 许可证

该项目基于 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件了解详情。