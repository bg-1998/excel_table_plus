import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:excel_table_plus/excel_table_plus.dart';

class JsonDemo extends StatefulWidget {
  const JsonDemo({super.key});

  @override
  State<JsonDemo> createState() => _JsonDemoState();
}

class _JsonDemoState extends State<JsonDemo> {
  late ExcelController _controller;
  String _jsonOutput = '';
  final TextEditingController _jsonInputController = TextEditingController();

  @override
  void initState() {
    super.initState();
    _initializeExcel();
  }

  void _initializeExcel() {
    // 初始化Excel模型
    final excelModel = ExcelModel(
      x: 8,
      y: 10,
      showSn: true,
      backgroundColor: Colors.white,
      rowColor: Colors.blue.withValues(alpha: 0.1),
      dividerWidth: 1.0,
      resizable: true,
      borderRadius: 8.0,
    );

    // 创建包含不同类型单元格的数据
    final items = <ExcelItemModel>[
      // 标题行
      ExcelItemModel(
        position: ExcelPosition(0, 0),
        value: TextCellValue(
          value: '姓名',
          style: const TextStyle(fontWeight: FontWeight.bold),
        ),
        cellType: 'text',
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      ExcelItemModel(
        position: ExcelPosition(1, 0),
        value: TextCellValue(
          value: '年龄',
          style: const TextStyle(fontWeight: FontWeight.bold),
        ),
        cellType: 'text',
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      ExcelItemModel(
        position: ExcelPosition(2, 0),
        value: TextCellValue(
          value: '邮箱',
          style: const TextStyle(fontWeight: FontWeight.bold),
        ),
        cellType: 'text',
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      ExcelItemModel(
        position: ExcelPosition(3, 0),
        value: TextCellValue(
          value: '部门',
          style: const TextStyle(fontWeight: FontWeight.bold),
        ),
        cellType: 'text',
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),

      // 数据行
      ExcelItemModel(
        position: ExcelPosition(0, 1),
        value: TextCellValue(value: '张三'),
        cellType: 'text',
      ),
      ExcelItemModel(
        position: ExcelPosition(1, 1),
        value: NumberCellValue(value: 28),
        cellType: 'number',
      ),
      ExcelItemModel(
        position: ExcelPosition(2, 1),
        value: TextCellValue(value: 'zhangsan@example.com'),
        cellType: 'text',
      ),
      ExcelItemModel(
        position: ExcelPosition(3, 1),
        value: TextCellValue(value: '工程部'),
        cellType: 'text',
      ),

      ExcelItemModel(
        position: ExcelPosition(0, 2),
        value: TextCellValue(value: '李四'),
        cellType: 'text',
      ),
      ExcelItemModel(
        position: ExcelPosition(1, 2),
        value: NumberCellValue(value: 32),
        cellType: 'number',
      ),
      ExcelItemModel(
        position: ExcelPosition(2, 2),
        value: TextCellValue(value: 'lisi@example.com'),
        cellType: 'text',
      ),
      ExcelItemModel(
        position: ExcelPosition(3, 2),
        value: TextCellValue(value: '设计部'),
        cellType: 'text',
      ),

      // 合并单元格示例
      ExcelItemModel(
        position: ExcelPosition(0, 4),
        value: TextCellValue(
          value: '合并单元格示例',
          style: const TextStyle(fontWeight: FontWeight.bold),
        ),
        cellType: 'text',
        color: Colors.green,
        isReadOnly: true,
        isMergeCell: true,
        positions: [
          ExcelPosition(0, 4),
          ExcelPosition(1, 4),
          ExcelPosition(2, 4),
          ExcelPosition(0, 5),
          ExcelPosition(1, 5),
          ExcelPosition(2, 5),
        ],
      ),
    ];

    // 初始化控制器
    _controller = ExcelController(
      excel: excelModel,
      items: items,
    );
  }

  // 导出表格数据为JSON
  void _exportToJson() {
    try {
      final jsonData = {
        'excel': _controller.excel.toJson(),
        'items': _controller.items.map((item) => item.toJson()).toList(),
      };

      final jsonString = JsonEncoder.withIndent('  ').convert(jsonData);
      setState(() {
        _jsonOutput = jsonString;
      });

      // 同时复制到剪贴板
      // Clipboard.setData(ClipboardData(text: jsonString));
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('数据已导出为JSON格式')),
      );
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('导出失败: $e')),
      );
    }
  }

  // 从JSON导入表格数据
  void _importFromJson() {
    try {
      if (_jsonInputController.text.isEmpty) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('请输入JSON数据')),
        );
        return;
      }

      final jsonData = json.decode(_jsonInputController.text) as Map<String, dynamic>;

      ExcelModel excel = ExcelModel.fromJson(jsonData['excel'] as Map<String, dynamic>);
      List<ExcelItemModel> items = (jsonData['items'] as List<dynamic>)
          .map((itemJson) => ExcelItemModel.fromJson(
                itemJson as Map<String, dynamic>,
                customValueFactory: (cellType, valueJson) {
                  // 根据cellType创建相应的自定义对象
                  switch (cellType) {
                    case 'text':
                      return TextCellValue.fromJson(valueJson);
                    case 'number':
                      return NumberCellValue.fromJson(valueJson);
                    default:
                      return valueJson;
                  }
                },
              ))
          .toList();

      setState(() {
        _controller = ExcelController(
          excel: excel,
          items: items,
        );
      });

      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('数据导入成功')),
      );
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('导入失败: $e')),
      );
    }
  }

  // 重置为初始状态
  void _resetToInitial() {
    setState(() {
      _initializeExcel();
      _jsonOutput = '';
      _jsonInputController.clear();
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('JSON导出/导入演示'),
        actions: [
          IconButton(
            icon: const Icon(Icons.refresh),
            onPressed: _resetToInitial,
            tooltip: '重置',
          ),
        ],
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              'JSON导出/导入功能演示',
              style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold),
            ),
            const SizedBox(height: 16),
            Expanded(
              flex: 3,
              child: ExcelWidget(
                controller: _controller,
                itemBuilder: (x, y, item) {
                  if (item != null) {
                    if (item.value is TextCellValue) {
                      final cellValue = item.value as TextCellValue;
                      return TextField(
                        controller: TextEditingController()..text = cellValue.value ?? '',
                        readOnly: item.isReadOnly,
                        decoration: const InputDecoration(
                          border: InputBorder.none,
                          contentPadding: EdgeInsets.all(8),
                        ),
                        style: cellValue.style,
                        onChanged: (value) {
                          cellValue.value = value;
                        },
                      );
                    } else if (item.value is NumberCellValue) {
                      final cellValue = item.value as NumberCellValue;
                      return TextField(
                        controller: TextEditingController()..text = cellValue.value?.toString() ?? '',
                        readOnly: item.isReadOnly,
                        decoration: const InputDecoration(
                          border: InputBorder.none,
                          contentPadding: EdgeInsets.all(8),
                        ),
                        keyboardType: TextInputType.number,
                        onChanged: (value) {
                          cellValue.value = int.tryParse(value);
                        },
                      );
                    }
                  }
                  return const SizedBox();
                },
              ),
            ),
            const SizedBox(height: 16),
            Expanded(
              flex: 2,
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(
                    children: [
                      ElevatedButton.icon(
                        onPressed: _exportToJson,
                        icon: const Icon(Icons.upload),
                        label: const Text('导出为JSON'),
                      ),
                      const SizedBox(width: 16),
                      ElevatedButton.icon(
                        onPressed: _importFromJson,
                        icon: const Icon(Icons.download),
                        label: const Text('从JSON导入'),
                      ),
                    ],
                  ),
                  const SizedBox(height: 16),
                  Expanded(
                    child: Row(
                      children: [
                        if (_jsonOutput.isNotEmpty) ...[
                          Expanded(
                            child: Column(
                              children: [
                                const Text(
                                  '导出的JSON数据:',
                                  style: TextStyle(fontWeight: FontWeight.bold),
                                ),
                                const SizedBox(height: 8),
                                Expanded(
                                  child: SingleChildScrollView(
                                    child: Container(
                                      padding: const EdgeInsets.all(8),
                                      decoration: BoxDecoration(
                                        border: Border.all(color: Colors.grey),
                                        borderRadius: BorderRadius.circular(4),
                                      ),
                                      child: SelectableText(
                                        _jsonOutput,
                                        style: const TextStyle(fontFamily: 'monospace', fontSize: 12),
                                      ),
                                    ),
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ] ,
                        Expanded(
                          child: Column(
                            children: [
                              const Text(
                                '导入JSON数据:',
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                              const SizedBox(height: 8),
                              Expanded(
                                child: TextField(
                                  controller: _jsonInputController,
                                  maxLines: null,
                                  expands: true,
                                  decoration: const InputDecoration(
                                    hintText: '在此粘贴JSON数据...',
                                    border: OutlineInputBorder(),
                                  ),
                                ),
                              ),
                            ],
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }
}

// 自定义文本单元格值类
class TextCellValue {
  String? value;
  final TextStyle? style;

  TextCellValue({this.value, this.style});

  Map<String, dynamic> toJson() {
    final Map<String, dynamic> json = <String, dynamic>{};
    json['value'] = value;
    if (style != null) {
      json['style'] = {
        'fontSize': style!.fontSize,
        'fontWeight': style!.fontWeight?.toString(),
        'color': style!.color?.value,
      };
    }
    return json;
  }

  factory TextCellValue.fromJson(Map<String, dynamic> json) {
    TextStyle? style;
    if (json['style'] != null) {
      final styleJson = json['style'] as Map<String, dynamic>;
      FontWeight? fontWeight;
      if (styleJson['fontWeight'] != null) {
        final fontWeightString = styleJson['fontWeight'] as String;
        switch (fontWeightString) {
          case 'FontWeight.bold':
            fontWeight = FontWeight.bold;
            break;
          default:
            fontWeight = FontWeight.normal;
        }
      }

      style = TextStyle(
        fontSize: styleJson['fontSize'] as double?,
        fontWeight: fontWeight,
        color: styleJson['color'] != null ? Color(styleJson['color'] as int) : null,
      );
    }

    return TextCellValue(
      value: json['value'] as String?,
      style: style,
    );
  }
}

// 自定义数字单元格值类
class NumberCellValue {
  int? value;

  NumberCellValue({this.value});

  Map<String, dynamic> toJson() {
    final Map<String, dynamic> json = <String, dynamic>{};
    json['value'] = value;
    return json;
  }

  factory NumberCellValue.fromJson(Map<String, dynamic> json) {
    return NumberCellValue(
      value: json['value'] as int?,
    );
  }
}