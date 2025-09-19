import 'package:flutter/material.dart';
import 'package:excel_table_plus/excel_table_plus.dart';

class TemplateDemo extends StatefulWidget {
  const TemplateDemo({super.key});

  @override
  State<TemplateDemo> createState() => _TemplateDemoState();
}

class _TemplateDemoState extends State<TemplateDemo> {
  late List<ExcelController> _controllers;
  final int x = 4;
  final int y = 3;

  @override
  void initState() {
    super.initState();
    _initializeTemplates();
  }

  void _initializeTemplates() {
    _controllers = [];

    // 模板1: 纯白色背景，无额外颜色
    _controllers.add(_createTemplate(
      templateName: '纯白风格',
      backgroundColor: Colors.white,
      rowColor: null,
      columnColor: null,
    ));

    // 模板2: 行交替颜色
    _controllers.add(_createTemplate(
      templateName: '行交替色',
      backgroundColor: Colors.white,
      rowColor: Colors.blue.withValues(alpha: 0.1),
      alternateRowColor: Colors.green.withValues(alpha: 0.1),
    ));

    // 模板3: 列交替颜色
    _controllers.add(_createTemplate(
      templateName: '列交替色',
      backgroundColor: Colors.white,
      columnColor: Colors.orange.withValues(alpha: 0.1),
      alternateColumnColor: Colors.purple.withValues(alpha: 0.1),
    ));

    // 模板4: 特殊表头行
    _controllers.add(_createTemplate(
      templateName: '彩色表头',
      backgroundColor: Colors.white,
      rowColor: Colors.blue.withValues(alpha: 0.05),
      headerColor: Colors.blue,
    ));

    // 模板5: 网格线更明显
    _controllers.add(_createTemplate(
      templateName: '明显网格',
      backgroundColor: Colors.white,
      rowColor: Colors.grey.withValues(alpha: 0.05),
      dividerWidth: 2.0,
      dividerColor: Colors.grey,
    ));

    // // 模板6: 圆角边框
    _controllers.add(_createTemplate(
      templateName: '圆角边框',
      backgroundColor: Colors.white,
      rowColor: Colors.deepPurple.withValues(alpha: 0.05),
      borderRadius: 12.0,
    ));
    
    // 模板7: 棋盘格样式
    _controllers.add(_createTemplate(
      templateName: '棋盘格',
      backgroundColor: Colors.white,
      checkerboardColor: Colors.blue.withValues(alpha: 0.1),
    ));
    
    // 模板8: 对角线样式
    _controllers.add(_createTemplate(
      templateName: '对角线',
      backgroundColor: Colors.white,
      diagonalColor: Colors.red.withValues(alpha: 0.2),
    ));
    
    // 模板9: 渐变色样式
    _controllers.add(_createTemplate(
      templateName: '渐变色',
      backgroundColor: Colors.white,
      gradientStartColor: Colors.yellow.withValues(alpha: 0.1),
      gradientEndColor: Colors.green.withValues(alpha: 0.3),
    ));
    
    // 模板10: 斑马纹（反向）
    _controllers.add(_createTemplate(
      templateName: '斑马纹（反向）',
      backgroundColor: Colors.white,
      rowColor: Colors.grey.withValues(alpha: 0.05),
      alternateRowColor: Colors.white,
    ));
    
    // // 模板11: 外边框高亮
    _controllers.add(_createTemplate(
      templateName: '外边框高亮',
      backgroundColor: Colors.white,
      borderColor: Colors.deepOrange.withValues(alpha: 0.5),
      borderWidth: 3.0,
    ));
    
    // 模板12: 随机色块
    _controllers.add(_createTemplate(
      templateName: '随机色块',
      backgroundColor: Colors.white,
      randomColor: true,
    ));
  }

  ExcelController _createTemplate({
    required String templateName,
    Color? backgroundColor,
    Color? rowColor,
    Color? columnColor,
    Color? alternateRowColor,
    Color? alternateColumnColor,
    Color? headerColor,
    Color? dividerColor,
    double? dividerWidth,
    double? borderRadius,
    Color? checkerboardColor,
    Color? diagonalColor,
    Color? gradientStartColor,
    Color? gradientEndColor,
    Color? borderColor,
    double? borderWidth,
    bool? randomColor,
  }) {
    final excelModel = ExcelModel(
      x: x,
      y: y,
      itemWidth: 90.0,
      itemHeight: 40.0,
      showSn: false,
      backgroundColor: backgroundColor ?? Colors.white,
      rowColor: rowColor,
      columnColor: columnColor,
      selectedBorderWidth: 2,
      dividerWidth: dividerWidth ?? 1.0,
      dividerColor: dividerColor ?? Colors.grey.withValues(alpha: 0.3),
      borderRadius: borderRadius ?? 0,
      resizable: false, // 模板不需要调整大小
      isEnableMultipleSelection: false,// 不允许多选
      borderWidth: borderWidth ?? 1,
      borderColor: borderColor,
    );

    // 创建单元格数据
    final items = <ExcelItemModel>[];

    // 如果有表头颜色，添加表头
    if (headerColor != null) {
      for (int i = 0; i < x; i++) {
        items.add(
          ExcelItemModel(
            position: ExcelPosition(i, 0),
            color: headerColor,
            isReadOnly: true,
          ),
        );
      }
    }

    // 添加行交替色
    if (alternateRowColor != null) {
      for (int row = 0; row < y; row++) {
        for (int col = 0; col < x; col++) {
          // 跳过表头行
          if (headerColor != null && row == 0) continue;
          
          // 偶数行使用交替颜色
          if (row % 2 == 0) {
            items.add(
              ExcelItemModel(
                position: ExcelPosition(col, row),
                color: alternateRowColor,
              ),
            );
          }
        }
      }
    }

    // 添加列交替色
    if (alternateColumnColor != null) {
      for (int row = 0; row < y; row++) {
        for (int col = 0; col < x; col++) {
          // 跳过表头行
          if (headerColor != null && row == 0) continue;
          
          // 偶数列使用交替颜色
          if (col % 2 == 0) {
            // 检查是否已在此位置添加了行颜色
            bool alreadyAdded = items.any((item) => 
              item.position.x == col && item.position.y == row);
            
            if (alreadyAdded) {
              // 如果已有颜色，混合颜色
              items.add(
                ExcelItemModel(
                  position: ExcelPosition(col, row),
                  color: alternateColumnColor.withValues(alpha: 0.3),
                ),
              );
            } else {
              items.add(
                ExcelItemModel(
                  position: ExcelPosition(col, row),
                  color: alternateColumnColor,
                ),
              );
            }
          }
        }
      }
    }
    
    // 添加棋盘格样式
    if (checkerboardColor != null) {
      for (int row = 0; row < y; row++) {
        for (int col = 0; col < x; col++) {
          // 跳过表头行
          if (headerColor != null && row == 0) continue;
          
          // 棋盘格样式：(行号+列号)为偶数时着色
          if ((row + col) % 2 == 0) {
            items.add(
              ExcelItemModel(
                position: ExcelPosition(col, row),
                color: checkerboardColor,
              ),
            );
          }
        }
      }
    }
    
    // 添加对角线样式
    if (diagonalColor != null) {
      for (int row = 0; row < y; row++) {
        for (int col = 0; col < x; col++) {
          // 跳过表头行
          if (headerColor != null && row == 0) continue;
          
          // 对角线样式：行号等于列号时着色
          if (row == col) {
            items.add(
              ExcelItemModel(
                position: ExcelPosition(col, row),
                color: diagonalColor,
              ),
            );
          }
        }
      }
    }
    
    // 添加渐变色样式
    if (gradientStartColor != null && gradientEndColor != null) {
      for (int row = 0; row < y; row++) {
        for (int col = 0; col < x; col++) {
          // 跳过表头行
          if (headerColor != null && row == 0) continue;
          
          // 计算渐变因子
          double factor = (row + col) / 7.0; // (0-3) + (0-4) = 0-7
          Color blendedColor = Color.lerp(gradientStartColor, gradientEndColor, factor)!;
          
          items.add(
            ExcelItemModel(
              position: ExcelPosition(col, row),
              color: blendedColor,
            ),
          );
        }
      }
    }
    
    // 添加随机色块样式
    if (randomColor == true) {
      final randomColors = [
        Colors.red.withValues(alpha: 0.1),
        Colors.blue.withValues(alpha: 0.1),
        Colors.green.withValues(alpha: 0.1),
        Colors.orange.withValues(alpha: 0.1),
        Colors.purple.withValues(alpha: 0.1),
      ];
      
      for (int row = 0; row < y; row++) {
        for (int col = 0; col < x; col++) {
          // 跳过表头行
          if (headerColor != null && row == 0) continue;
          
          // 每隔一个单元格应用随机颜色
          if ((row + col) % 2 == 1) {
            final randomColor = randomColors[(row * x + col) % randomColors.length];
            items.add(
              ExcelItemModel(
                position: ExcelPosition(col, row),
                color: randomColor,
              ),
            );
          }
        }
      }
    }

    return ExcelController(
      excel: excelModel,
      items: items,
      isAutoDispose: false,
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('表格模板样式演示'),
      ),
      body: GridView.builder(
        gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(
          crossAxisCount: 2,
          childAspectRatio: 2,
          crossAxisSpacing: 16,
          mainAxisSpacing: 16,
        ),
        padding: const EdgeInsets.all(20),
        itemCount: _controllers.length,
        itemBuilder: (context, index) {
              double itemHeight = 40.0;
              double borderWidth = _controllers[index].excel.borderWidth;
              double dividerWidth = _controllers[index].excel.dividerWidth;
              int rowCount = _controllers[index].excel.y;
              return FittedBox(
                child: SizedBox(
                  height: (itemHeight)*rowCount+dividerWidth*(rowCount-1)+40+borderWidth*2,
                  child: Column(
                    children: [
                      Container(
                        height: 40,
                        alignment: Alignment.center,
                        child: Text(
                          _getTemplateName(index),
                          style: const TextStyle(
                            fontWeight: FontWeight.bold,
                            fontSize: 16,
                          ),
                        ),
                      ),
                      ExcelWidget(
                        controller: _controllers[index],
                        itemBuilder: (x, y, item) {
                          return const SizedBox();
                        },
                      ),
                    ],
                  )),
              );
        },
      ),
    );
  }

  String _getTemplateName(int index) {
    switch (index) {
      case 0:
        return '纯白风格';
      case 1:
        return '行交替色';
      case 2:
        return '列交替色';
      case 3:
        return '彩色表头';
      case 4:
        return '明显网格';
      case 5:
        return '圆角边框';
      case 6:
        return '棋盘格';
      case 7:
        return '对角线';
      case 8:
        return '渐变色';
      case 9:
        return '斑马纹（反向）';
      case 10:
        return '外边框高亮';
      case 11:
        return '随机色块';
      default:
        return '模板 ${index + 1}';
    }
  }
}