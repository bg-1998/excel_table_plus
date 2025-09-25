import 'package:excel_table_plus/excel_table_plus.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

import 'resize_handle_painter.dart';

// 定义不同的单元格类型
enum CellType {
  text,
  number,
  date,
  dropdown,
}

class AdvancedExcelDemo extends StatefulWidget {
  const AdvancedExcelDemo({super.key});

  @override
  State<AdvancedExcelDemo> createState() => _AdvancedExcelDemoState();
}

class _AdvancedExcelDemoState extends State<AdvancedExcelDemo> {
  late ExcelController _controller;
  double _scale = 1.0;
  bool _isResizing = false;
  Offset _lastFocalPoint = Offset.zero;
  bool _showToolbar = false;

  @override
  void initState() {
    super.initState();
    _initializeExcel();
  }

  void _initializeExcel() {
    // 创建Excel模型
    final excelModel = ExcelModel(
      x: 15, // 15列
      y: 30, // 30行
      showSn: true, // 显示序号
      backgroundColor: Colors.white,
      rowColor: Colors.indigo.withValues(alpha: 0.05), // 行背景色
      columnColor: Colors.indigo.withValues(alpha: 0.02), // 列背景色
      dividerWidth: 0.5, // 分割线宽度
      dividerColor: Colors.grey.withValues(alpha: 0.3),
      borderRadiusTL: 8.0,
      borderRadiusTR: 8.0,
      borderRadiusBL: 8.0,
      borderRadiusBR: 8.0,
      resizable: true, // 允许调整大小
      resizeAreaSize: 8.0,
      alignment: Alignment.center, // 居中对齐
      // 自定义列宽
      customColumnWidths: [
        100, 120, 120, 120, 120, 150, 100, 100, 100, 100, 100, 100, 100, 100, 100
      ],
      // 自定义行高
      customRowHeights: List.generate(30, (index) => index == 0 ? 50.0 : 40.0),
    );

    // 创建单元格数据
    final items = <ExcelItemModel>[];

    // 添加表头
    final headers = [
      '项目名称', '开始日期', '结束日期', '负责人', '状态', '备注',
      '预算', '实际支出', '完成度', '优先级', '风险等级', '部门', '类型', '阶段', '评分'
    ];

    for (int i = 0; i < headers.length; i++) {
      items.add(
        ExcelItemModel(
          position: ExcelPosition(i, 0),
          value: AdvancedCellValue(
            value: headers[i],
            cellType: CellType.text,
            style: const TextStyle(
              fontWeight: FontWeight.bold,
              fontSize: 14,
            ),
          ),
          color: Colors.indigo.withValues(alpha: 0.1),
          isReadOnly: true,
        ),
      );
    }

    // 添加示例数据
    final projectNames = ['产品设计', '前端开发', '后端开发', '测试', '部署上线', '用户反馈', '迭代优化'];
    final statuses = ['未开始', '进行中', '已完成', '已延期', '已取消'];
    final departments = ['研发部', '设计部', '测试部', '运维部', '市场部'];
    final types = ['功能开发', '缺陷修复', '性能优化', '安全加固', '体验提升'];
    final stages = ['规划', '开发', '测试', '发布', '维护'];

    for (int row = 1; row < 10; row++) {
      final projectIndex = (row - 1) % projectNames.length;
      final statusIndex = (row - 1) % statuses.length;

      // 项目名称
      items.add(
        ExcelItemModel(
          position: ExcelPosition(0, row),
          value: AdvancedCellValue(
            value: projectNames[projectIndex],
            cellType: CellType.text,
          ),
        ),
      );

      // 开始日期
      items.add(
        ExcelItemModel(
          position: ExcelPosition(1, row),
          value: AdvancedCellValue(
            value: '2023-${(row % 12) + 1}-${(row % 28) + 1}',
            cellType: CellType.date,
          ),
        ),
      );

      // 结束日期
      items.add(
        ExcelItemModel(
          position: ExcelPosition(2, row),
          value: AdvancedCellValue(
            value: '2023-${((row + 2) % 12) + 1}-${((row + 15) % 28) + 1}',
            cellType: CellType.date,
          ),
        ),
      );

      // 负责人
      items.add(
        ExcelItemModel(
          position: ExcelPosition(3, row),
          value: AdvancedCellValue(
            value: '负责人${row}',
            cellType: CellType.text,
          ),
        ),
      );

      // 状态
      items.add(
        ExcelItemModel(
          position: ExcelPosition(4, row),
          value: AdvancedCellValue(
            value: statuses[statusIndex],
            cellType: CellType.dropdown,
            dropdownItems: statuses,
          ),
          color: _getStatusColor(statuses[statusIndex]),
        ),
      );

      // 备注
      items.add(
        ExcelItemModel(
          position: ExcelPosition(5, row),
          value: AdvancedCellValue(
            value: '这是项目${row}的备注信息，可以包含多行内容。',
            cellType: CellType.text,
          ),
        ),
      );

      // 预算
      items.add(
        ExcelItemModel(
          position: ExcelPosition(6, row),
          value: AdvancedCellValue(
            value: '${(row * 10000).toString()}',
            cellType: CellType.number,
            textAlign: TextAlign.right,
          ),
        ),
      );

      // 实际支出
      items.add(
        ExcelItemModel(
          position: ExcelPosition(7, row),
          value: AdvancedCellValue(
            value: '${(row * 8500).toString()}',
            cellType: CellType.number,
            textAlign: TextAlign.right,
          ),
        ),
      );

      // 完成度
      items.add(
        ExcelItemModel(
          position: ExcelPosition(8, row),
          value: AdvancedCellValue(
            value: '${(row * 10) % 100}%',
            cellType: CellType.text,
            textAlign: TextAlign.center,
          ),
        ),
      );

      // 优先级
      items.add(
        ExcelItemModel(
          position: ExcelPosition(9, row),
          value: AdvancedCellValue(
            value: '${(row % 3) + 1}',
            cellType: CellType.number,
            textAlign: TextAlign.center,
          ),
        ),
      );

      // 风险等级
      items.add(
        ExcelItemModel(
          position: ExcelPosition(10, row),
          value: AdvancedCellValue(
            value: '${(row % 5) + 1}',
            cellType: CellType.number,
            textAlign: TextAlign.center,
          ),
        ),
      );

      // 部门
      items.add(
        ExcelItemModel(
          position: ExcelPosition(11, row),
          value: AdvancedCellValue(
            value: departments[row % departments.length],
            cellType: CellType.dropdown,
            dropdownItems: departments,
          ),
        ),
      );

      // 类型
      items.add(
        ExcelItemModel(
          position: ExcelPosition(12, row),
          value: AdvancedCellValue(
            value: types[row % types.length],
            cellType: CellType.dropdown,
            dropdownItems: types,
          ),
        ),
      );

      // 阶段
      items.add(
        ExcelItemModel(
          position: ExcelPosition(13, row),
          value: AdvancedCellValue(
            value: stages[row % stages.length],
            cellType: CellType.dropdown,
            dropdownItems: stages,
          ),
        ),
      );

      // 评分
      items.add(
        ExcelItemModel(
          position: ExcelPosition(14, row),
          value: AdvancedCellValue(
            value: '${(row % 5) + 1}',
            cellType: CellType.number,
            textAlign: TextAlign.center,
          ),
        ),
      );
    }

    // 添加合并单元格示例
    items.add(
      ExcelItemModel(
        position: ExcelPosition(0, 15),
        value: AdvancedCellValue(
          value: '合并单元格示例 - 项目总结',
          cellType: CellType.text,
          style: const TextStyle(
            fontWeight: FontWeight.bold,
            fontSize: 16,
          ),
          textAlign: TextAlign.center,
        ),
        color: Colors.amber,
        isReadOnly: true,
        isMergeCell: true,
        positions: [
          ExcelPosition(0, 15),
          ExcelPosition(1, 15),
          ExcelPosition(2, 15),
          ExcelPosition(3, 15),
          ExcelPosition(4, 15),
        ],
      ),
    );

    // 初始化控制器
    _controller = ExcelController(
      excel: excelModel,
      items: items,
    );

    // 设置单元格选中回调
    _controller.onPositionSelected = (position) {
      setState(() {
        _showToolbar = position != null;
      });
    };

    // 设置多选回调
    _controller.onMultiSelectionChanged = (selectedItems) {
      setState(() {
        _showToolbar = selectedItems.isNotEmpty;
      });
    };
  }

  // 根据状态获取颜色
  Color _getStatusColor(String status) {
    switch (status) {
      case '未开始':
        return Colors.grey.withValues(alpha: 0.2);
      case '进行中':
        return Colors.blue.withValues(alpha: 0.2);
      case '已完成':
        return Colors.green.withValues(alpha: 0.2);
      case '已延期':
        return Colors.orange.withValues(alpha: 0.2);
      case '已取消':
        return Colors.red.withValues(alpha: 0.2);
      default:
        return Colors.transparent;
    }
  }

  // 更新缩放比例
  void _updateScale(double newScale) {
    _controller.excel.itemWidth = _controller.startScaleExcel!.itemWidth * newScale;
    _controller.excel.itemHeight = _controller.startScaleExcel!.itemHeight * newScale;
    _controller.excel.dividerWidth = _controller.startScaleExcel!.dividerWidth * newScale;
    _controller.excel.borderRadiusTL = _controller.startScaleExcel!.borderRadiusTL * newScale;
    _controller.excel.borderRadiusTR = _controller.startScaleExcel!.borderRadiusTR * newScale;
    _controller.excel.borderRadiusBL = _controller.startScaleExcel!.borderRadiusBL * newScale;
    _controller.excel.borderRadiusBR = _controller.startScaleExcel!.borderRadiusBR * newScale;
    _controller.excel.resizeAreaSize = _controller.startScaleExcel!.resizeAreaSize * newScale;
    _controller.excel.customColumnWidths = _controller.startScaleExcel!.customColumnWidths.map((width) => width * newScale)
        .toList();
    _controller.excel.customRowHeights = _controller.startScaleExcel!.customRowHeights.map((height) => height * newScale)
        .toList();
    // 更新单元格字体大小
    for(int i = 0; i < _controller.items.length; i++){
      if(_controller.items[i].value is AdvancedCellValue){
        var cell = _controller.items[i].value as AdvancedCellValue;
        var startCell = _controller.startScaleItems![i].value as AdvancedCellValue;
        cell.style = startCell.style.copyWith(
          fontSize: startCell.style.fontSize! * newScale,
        );
      }
    }
    _controller.update();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('高级Excel表格示例'),
        actions: [
          IconButton(
            icon: const Icon(Icons.refresh),
            tooltip: '重置',
            onPressed: () {
              setState(() {
                _initializeExcel();
                _scale = 1.0;
              });
            },
          ),
          SizedBox(width: 50,),
        ],
      ),
      body: Column(
        children: [
          // Excel表格
          Expanded(
            child: Padding(
              padding: const EdgeInsets.all(16.0),
              child: LayoutBuilder(builder: (context,constraints){
                  return SizedBox(
                    height: constraints.maxHeight,
                    width: constraints.maxWidth,
                    child: Stack(
                      children: [
                        ExcelWidget(
                          controller: _controller,
                          itemBuilder: (x, y, item) {
                            if (item != null && item.value is AdvancedCellValue) {
                              final cellValue = item.value as AdvancedCellValue;
                              cellValue.controller.text = cellValue.value ?? '';

                              // 根据单元格类型构建不同的编辑器
                              switch (cellValue.cellType) {
                                case CellType.dropdown:
                                  return _buildDropdownCell(item, cellValue);
                                case CellType.date:
                                  return _buildDateCell(item, cellValue);
                                case CellType.number:
                                  return _buildNumberCell(item, cellValue);
                                case CellType.text:
                                default:
                                  return _buildTextCell(item, cellValue);
                              }
                            } else {
                              return const SizedBox();
                            }
                          },
                        ),
                        Positioned(
                          right: 0,
                          bottom: 0.0,
                          child: MouseRegion(
                            cursor: SystemMouseCursors.resizeRight,
                            child: GestureDetector(
                              onPanStart: (details) {
                                setState(() {
                                  _scale = 1;
                                  _controller.startScaleExcel = _controller.excel;
                                  _controller.startScaleItems = _controller.items;
                                  _isResizing = true;
                                  _lastFocalPoint = details.globalPosition;
                                });
                              },
                              onPanUpdate: (details) {
                                if (!_isResizing) return;

                                // 计算手势移动距离
                                final delta = details.globalPosition - _lastFocalPoint;
                                _lastFocalPoint = details.globalPosition;

                                // 根据移动距离调整缩放比例
                                // 这里使用一个简单的计算方式，你可以根据需要调整灵敏度
                                final scaleFactor = 0.01;
                                final scaleDelta = (delta.dx + delta.dy) * scaleFactor;
                                final newScale = _scale + scaleDelta;

                                _updateScale(newScale);
                              },
                              onPanEnd: (details) {
                                setState(() {
                                  _isResizing = false;
                                });
                              },
                              child: Container(
                                width: 20,
                                height: 20,
                                color: _isResizing ? Colors.blue.withValues(alpha: 0.3) : Colors.transparent,
                                child: CustomPaint(
                                  painter: ResizeHandlePainter(),
                                ),
                              ),
                            ),
                          ),
                        ),
                      ],
                    ),
                  );
                }
              ),
            ),
          ),
        ],
      ),
      // 底部工具栏
      bottomNavigationBar: _showToolbar
          ? BottomAppBar(
              child: Container(
                height: 60,
                padding: const EdgeInsets.symmetric(horizontal: 16.0),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    _buildToolbarButton(
                      icon: Icons.merge_type,
                      label: '合并',
                      onPressed: () => _controller.mergeSelectedCells(),
                    ),
                    _buildToolbarButton(
                      icon: Icons.call_split,
                      label: '拆分',
                      onPressed: () => _controller.splitMergedCell(),
                    ),
                    _buildToolbarButton(
                      icon: Icons.add_box,
                      label: '插行',
                      onPressed: _insertRow,
                    ),
                    _buildToolbarButton(
                      icon: Icons.add_box_outlined,
                      label: '插列',
                      onPressed: _insertColumn,
                    ),
                    _buildToolbarButton(
                      icon: Icons.delete,
                      label: '删行',
                      onPressed: _deleteRow,
                    ),
                    _buildToolbarButton(
                      icon: Icons.delete_outline,
                      label: '删列',
                      onPressed: _deleteColumn,
                    ),
                    _buildToolbarButton(
                      icon: Icons.format_color_fill,
                      label: '背景',
                      onPressed: _setBackgroundColor,
                    ),
                    _buildToolbarButton(
                      icon: Icons.format_bold,
                      label: '加粗',
                      onPressed: _setBold,
                    ),
                  ],
                ),
              ),
            )
          : SizedBox(height: 80,),
    );
  }

  // 构建文本单元格
  Widget _buildTextCell(ExcelItemModel item, AdvancedCellValue cellValue) {
    return TextField(
      controller: cellValue.controller,
      textAlign: cellValue.textAlign,
      style: cellValue.style,
      readOnly: item.isReadOnly,
      decoration: const InputDecoration(
        border: InputBorder.none,
        isCollapsed: true,
      ),
      onChanged: (value) {
        cellValue.value = value;
      },
    );
  }

  // 构建数字单元格
  Widget _buildNumberCell(ExcelItemModel item, AdvancedCellValue cellValue) {
    return TextField(
      controller: cellValue.controller,
      textAlign: cellValue.textAlign,
      style: cellValue.style,
      readOnly: item.isReadOnly,
      keyboardType: const TextInputType.numberWithOptions(decimal: true),
      inputFormatters: [
        FilteringTextInputFormatter.allow(RegExp(r'[0-9.]')),
      ],
      decoration: const InputDecoration(
        border: InputBorder.none,
        contentPadding: EdgeInsets.symmetric(horizontal: 8, vertical: 4),
        isCollapsed: true,
      ),
      onChanged: (value) {
        cellValue.value = value;
      },
    );
  }

  // 构建日期单元格
  Widget _buildDateCell(ExcelItemModel item, AdvancedCellValue cellValue) {
    return GestureDetector(
      onTap: !item.isReadOnly
          ? () async {
              final date = await showDatePicker(
                context: context,
                initialDate: DateTime.now(),
                firstDate: DateTime(2000),
                lastDate: DateTime(2100),
              );

              if (date != null) {
                final formattedDate =
                    '${date.year}-${date.month.toString().padLeft(2, '0')}-${date.day.toString().padLeft(2, '0')}';
                setState(() {
                  cellValue.value = formattedDate;
                  cellValue.controller.text = formattedDate;
                });
              }
            }
          : null,
      child: LayoutBuilder(
        builder: (context,constraints) {
          return FittedBox(
            child: Padding(
              padding: const EdgeInsets.symmetric(horizontal: 8.0),
              child: ConstrainedBox(
                constraints: BoxConstraints(
                  minWidth: 48,minHeight: 16,
                ),
                child: Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    Text(cellValue.controller.text,textAlign: cellValue.textAlign,
                      style: cellValue.style,),
                    SizedBox(width: 20,),
                    Icon(Icons.calendar_today, size: 12),
                  ],
                ),
              ),
            ),
          );
        }
      ),
    );
  }

  // 构建下拉单元格
  Widget _buildDropdownCell(ExcelItemModel item, AdvancedCellValue cellValue) {
    return GestureDetector(
      onTap: !item.isReadOnly
          ? () async {
              final result = await showModalBottomSheet<String>(
                context: context,
                builder: (BuildContext context) {
                  return SafeArea(
                    child: Column(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        ListTile(
                          title: const Text('选择选项'),
                          trailing: IconButton(
                            icon: const Icon(Icons.close),
                            onPressed: () => Navigator.of(context).pop(),
                          ),
                        ),
                        const Divider(),
                        ...(cellValue.dropdownItems?.map((String value) {
                              return ListTile(
                                title: Text(value),
                                onTap: () => Navigator.of(context).pop(value),
                              );
                            }).toList() ??
                            []),
                      ],
                    ),
                  );
                },
              );

              if (result != null) {
                setState(() {
                  cellValue.value = result;
                  cellValue.controller.text = result;

                  // 更新单元格背景色（如果是状态列）
                  if (item.position.x == 4) {
                    item.color = _getStatusColor(result);
                  }
                });
              }
            }
          : null,
      child: AbsorbPointer(
        child: LayoutBuilder(
            builder: (context,constraints) {
            return FittedBox(
              child: Padding(
                padding: const EdgeInsets.symmetric(horizontal: 8.0),
                child: ConstrainedBox(
                  constraints: BoxConstraints(
                    minWidth: 48,minHeight: 20,
                  ),
                  child: Row(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      Text(cellValue.controller.text,textAlign: cellValue.textAlign,
                        style: cellValue.style,
                      ),
                      SizedBox(width: 20,),
                      Icon(Icons.arrow_drop_down, size: 16),
                    ],
                  ),
                ),
              ),
            );
          }
        ),
      ),
    );
  }

  // 构建工具栏按钮
  Widget _buildToolbarButton({
    required IconData icon,
    required String label,
    required VoidCallback onPressed,
  }) {
    return InkWell(
      onTap: onPressed,
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Icon(icon, size: 20),
          const SizedBox(height: 4),
          Text(label, style: const TextStyle(fontSize: 12)),
        ],
      ),
    );
  }

  // 插入行
  void _insertRow() {
    if (_controller.selectedPosition != null) {
      _controller.insertRow(_controller.selectedPosition!.y);
      setState(() {});
    } else if (_controller.selectedItems.isNotEmpty) {
      _controller.selectedItems.sort((ExcelPosition a, ExcelPosition b) => a.y.compareTo(b.y));
      _controller.insertRow(_controller.selectedItems.first.y);
      setState(() {});
    }
  }

  // 插入列
  void _insertColumn() {
    if (_controller.selectedPosition != null) {
      _controller.insertColumn(_controller.selectedPosition!.x);
      setState(() {});
    } else if (_controller.selectedItems.isNotEmpty) {
      _controller.selectedItems.sort((ExcelPosition a, ExcelPosition b) => a.x.compareTo(b.x));
      _controller.insertColumn(_controller.selectedItems.first.x);
      setState(() {});
    }
  }

  // 删除行
  void _deleteRow() {
    if (_controller.selectedPosition != null) {
      _controller.deleteRow(_controller.selectedPosition!.y);
      setState(() {});
    } else if (_controller.selectedItems.isNotEmpty) {
      _controller.selectedItems.sort((ExcelPosition a, ExcelPosition b) => a.y.compareTo(b.y));
      _controller.deleteRow(_controller.selectedItems.first.y);
      setState(() {});
    }
  }

  // 删除列
  void _deleteColumn() {
    if (_controller.selectedPosition != null) {
      _controller.deleteColumn(_controller.selectedPosition!.x);
      setState(() {});
    } else if (_controller.selectedItems.isNotEmpty) {
      _controller.selectedItems.sort((ExcelPosition a, ExcelPosition b) => a.x.compareTo(b.x));
      _controller.deleteColumn(_controller.selectedItems.first.x);
      setState(() {});
    }
  }

  // 设置背景色
  void _setBackgroundColor() {
    // 简化版的颜色选择
    final colors = [
      Colors.transparent,
      Colors.red,
      Colors.orange,
      Colors.yellow,
      Colors.green,
      Colors.blue,
      Colors.indigo,
      Colors.purple,
      Colors.grey,
    ];

    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('选择背景色'),
        content: Wrap(
          spacing: 8,
          runSpacing: 8,
          children: colors.map((color) {
            return GestureDetector(
              onTap: () {
                Navigator.of(context).pop();
                _applyBackgroundColor(color);
              },
              child: Container(
                width: 40,
                height: 40,
                decoration: BoxDecoration(
                  color: color,
                  border: Border.all(color: Colors.grey),
                  borderRadius: BorderRadius.circular(4),
                ),
              ),
            );
          }).toList(),
        ),
      ),
    );
  }

  // 应用背景色
  void _applyBackgroundColor(Color color) {
    if (_controller.selectedItems.isNotEmpty) {
      for (var position in _controller.selectedItems) {
        var item = _controller.findItemContainingPosition(position.x, position.y);
        if (item != null) {
          item.color = color;
        } else {
          // 创建新的单元格
          _controller.items.add(
            ExcelItemModel(
              position: position,
              color: color,
            ),
          );
        }
      }
    } else if (_controller.selectedPosition != null) {
      var item = _controller.findItemContainingPosition(
        _controller.selectedPosition!.x,
        _controller.selectedPosition!.y,
      );
      if (item != null) {
        item.color = color;
      } else {
        // 创建新的单元格
        _controller.items.add(
          ExcelItemModel(
            position: _controller.selectedPosition!,
            color: color,
          ),
        );
      }
    }
    setState(() {});
  }

  // 设置文字加粗
  void _setBold() {
    if (_controller.selectedItems.isNotEmpty) {
      for (var position in _controller.selectedItems) {
        _toggleBoldForCell(position.x, position.y);
      }
    } else if (_controller.selectedPosition != null) {
      _toggleBoldForCell(
        _controller.selectedPosition!.x,
        _controller.selectedPosition!.y,
      );
    }
    setState(() {});
  }

  // 切换单元格文字加粗状态
  void _toggleBoldForCell(int x, int y) {
    var item = _controller.findItemContainingPosition(x, y);
    if (item != null && item.value is AdvancedCellValue) {
      final cellValue = item.value as AdvancedCellValue;
      final currentStyle = cellValue.style;
      final isBold = currentStyle.fontWeight == FontWeight.bold;

      cellValue.style = currentStyle.copyWith(
        fontWeight: isBold ? FontWeight.normal : FontWeight.bold,
      );
    }
  }
}

// 高级单元格值类
class AdvancedCellValue {
  String? value;
  CellType cellType;
  final TextEditingController controller;
  TextAlign textAlign;
  TextStyle style;
  List<String>? dropdownItems;

  AdvancedCellValue({
    this.value,
    required this.cellType,
    TextEditingController? controller,
    this.textAlign = TextAlign.center,
    TextStyle? style,
    this.dropdownItems,
  }) : controller = controller ?? TextEditingController(),
        style = style ?? TextStyle(fontSize: 14);
}