import 'package:flutter/material.dart';
import 'package:excel_table_plus/excel_table_plus.dart';

import 'resize_handle_painter.dart';

class ExcelTableDemo extends StatefulWidget {
  const ExcelTableDemo({super.key});

  @override
  State<ExcelTableDemo> createState() => _ExcelTableDemoState();
}

class _ExcelTableDemoState extends State<ExcelTableDemo> {
  // 创建Excel控制器
  late ExcelController _controller;
  bool _showToolbar = false;
  double _scale = 1.0;
  bool _isResizing = false;
  Offset _lastFocalPoint = Offset.zero;
  
  @override
  void initState() {
    super.initState();
    _initializeExcel();
  }
  
  void _initializeExcel() {
    // 初始化Excel模型
    final excelModel = ExcelModel(
      x: 10, // 10列
      y: 20, // 20行
      showSn: true, // 显示序号
      backgroundColor: Colors.white,
      rowColor: Colors.blue.withValues(alpha: 0.1), // 行背景色
      dividerWidth: 1, // 分割线宽度
      resizable: true, // 允许调整大小
      borderRadiusTL: 8,
      borderRadiusTR: 8,
      borderRadiusBL: 8,
      borderRadiusBR: 8,
    );
    
    // 创建单元格数据
    final items = <ExcelItemModel>[
      // 标题行
      ExcelItemModel(
        position: ExcelPosition(0, 0),
        value: _ExcelCellValue(value: '姓名'),
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      ExcelItemModel(
        position: ExcelPosition(1, 0),
        value: _ExcelCellValue(value: '年龄'),
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      ExcelItemModel(
        position: ExcelPosition(2, 0),
        value: _ExcelCellValue(value: '性别'),
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      ExcelItemModel(
        position: ExcelPosition(3, 0),
        value: _ExcelCellValue(value: '职业'),
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      ExcelItemModel(
        position: ExcelPosition(4, 0),
        value: _ExcelCellValue(value: '备注'),
        color: Colors.grey.shade200,
        isReadOnly: true,
      ),
      
      // 数据行1
      ExcelItemModel(
        position: ExcelPosition(0, 1),
        value: _ExcelCellValue(value: '张三'),
      ),
      ExcelItemModel(
        position: ExcelPosition(1, 1),
        value: _ExcelCellValue(value: '28'),
      ),
      ExcelItemModel(
        position: ExcelPosition(2, 1),
        value: _ExcelCellValue(value: '男'),
      ),
      ExcelItemModel(
        position: ExcelPosition(3, 1),
        value: _ExcelCellValue(value: '工程师'),
      ),
      ExcelItemModel(
        position: ExcelPosition(4, 1),
        value: _ExcelCellValue(value: '前端开发工程师'),
      ),
      
      // 数据行2
      ExcelItemModel(
        position: ExcelPosition(0, 2),
        value: _ExcelCellValue(value: '李四'),
      ),
      ExcelItemModel(
        position: ExcelPosition(1, 2),
        value: _ExcelCellValue(value: '32'),
      ),
      ExcelItemModel(
        position: ExcelPosition(2, 2),
        value: _ExcelCellValue(value: '男'),
      ),
      ExcelItemModel(
        position: ExcelPosition(3, 2),
        value: _ExcelCellValue(value: '设计师'),
      ),
      ExcelItemModel(
        position: ExcelPosition(4, 2),
        value: _ExcelCellValue(value: 'UI设计师'),
      ),
      
      // 数据行3
      ExcelItemModel(
        position: ExcelPosition(0, 3),
        value: _ExcelCellValue(value: '王五'),
      ),
      ExcelItemModel(
        position: ExcelPosition(1, 3),
        value: _ExcelCellValue(value: '25'),
      ),
      ExcelItemModel(
        position: ExcelPosition(2, 3),
        value: _ExcelCellValue(value: '女'),
      ),
      ExcelItemModel(
        position: ExcelPosition(3, 3),
        value: _ExcelCellValue(value: '产品经理'),
      ),
      ExcelItemModel(
        position: ExcelPosition(4, 3),
        value: _ExcelCellValue(value: '负责产品规划'),
      ),
      
      // 合并单元格示例
      ExcelItemModel(
        position: ExcelPosition(0, 5),
        value: _ExcelCellValue(value: '合并单元格示例'),
        color: Colors.green.withValues(alpha: 0.2),
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
    ];
    
    // 初始化控制器
    _controller = ExcelController(
      excel: excelModel,
      items: items,
    );
    
    // 设置单元格选中回调
    _controller.onPositionSelected = (ExcelPosition? position) {
      setState(() {
        _showToolbar = position != null;
      });
      if (position != null) {
        print('选中单元格: (${position.x}, ${position.y})');
      }
    };
    
    // 设置多选回调
    _controller.onMultiSelectionChanged = (selectedItems) {
      setState(() {
        _showToolbar = selectedItems.isNotEmpty;
      });
      if (selectedItems.isNotEmpty) {
        print('多选了 ${selectedItems.length} 个单元格');
      }
    };
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
      if(_controller.items[i].value is _ExcelCellValue){
        var cell = _controller.items[i].value as _ExcelCellValue;
        var startCell = _controller.startScaleItems![i].value as _ExcelCellValue;
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
        title: const Text('Excel Table Demo'),
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
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              '基本Excel表格示例',
              style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold),
            ),
            const SizedBox(height: 8),
            Expanded(
              child: LayoutBuilder(builder: (context,constraints){
                return Stack(
                  children: [
                    SizedBox(
                      width: constraints.maxWidth,
                      height: constraints.maxHeight,
                      child: ExcelWidget(
                        controller: _controller,
                        itemBuilder: (x, y, item) {
                          if (item != null && item.value is _ExcelCellValue) {
                            final cellValue = item.value as _ExcelCellValue;

                            // 创建文本编辑控制器
                            cellValue.controller.text = cellValue.value ?? '';

                            return TextField(
                              controller: cellValue.controller,
                              textAlign: cellValue.textAlign,
                              textAlignVertical: TextAlignVertical.top,
                              readOnly: item.isReadOnly,
                              decoration: const InputDecoration(
                                border: InputBorder.none,
                                isCollapsed: true,
                              ),
                              style: cellValue.style,
                              onChanged: (value) {
                                cellValue.value = value;
                              },
                            );
                          } else {
                            return const SizedBox();
                          }
                        },
                      ),
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
                );
              }),
            ),
          ],
        ),
      ),
      // 底部工具栏
      bottomNavigationBar: _showToolbar
          ? BottomAppBar(
              child: Padding(
                padding: const EdgeInsets.symmetric(horizontal: 16.0),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    _buildActionButton(
                      icon: Icons.merge_type,
                      label: '合并',
                      onPressed: () => _controller.mergeSelectedCells(),
                    ),
                    _buildActionButton(
                      icon: Icons.call_split,
                      label: '拆分',
                      onPressed: () => _controller.splitMergedCell(),
                    ),
                    _buildActionButton(
                      icon: Icons.add,
                      label: '插行',
                      onPressed: _insertRow,
                    ),
                    _buildActionButton(
                      icon: Icons.add_box,
                      label: '插列',
                      onPressed: _insertColumn,
                    ),
                    _buildActionButton(
                      icon: Icons.format_color_fill,
                      label: '背景',
                      onPressed: _setBackgroundColor,
                    ),
                  ],
                ),
              ),
            )
          : SizedBox(height: 80,),
    );
  }
  
  Widget _buildActionButton({
    required IconData icon,
    required String label,
    required VoidCallback onPressed,
  }) {
    return InkWell(
      onTap: onPressed,
      child: Padding(
        padding: const EdgeInsets.symmetric(vertical: 8.0),
        child: FittedBox(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Icon(icon, size: 20),
              const SizedBox(height: 4),
              Text(label, style: const TextStyle(fontSize: 12)),
            ],
          ),
        ),
      ),
    );
  }
  
  // 插入行
  void _insertRow() {
    if (_controller.selectedPosition != null) {
      _controller.insertRow(_controller.selectedPosition!.y);
      setState(() {});
    } else if (_controller.selectedItems.isNotEmpty) {
      // 如果有多个选中的单元格，则在第一个选中单元格所在行的上方插入新行
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
      // 如果有多个选中的单元格，则在第一个选中单元格所在列的左侧插入新列
      _controller.selectedItems.sort((ExcelPosition a, ExcelPosition b) => a.x.compareTo(b.x));
      _controller.insertColumn(_controller.selectedItems.first.x);
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
}

// 自定义单元格值类
class _ExcelCellValue {
  String? value;
  final TextEditingController controller;
  TextAlign textAlign;
  TextStyle style;
  
  _ExcelCellValue({
    this.value,
    TextEditingController? controller,
    TextAlign? textAlign,
    TextStyle? style,
  }) : controller = controller ?? TextEditingController(),
        textAlign = textAlign ?? TextAlign.center,
        style = style ?? TextStyle(fontSize: 14);
}