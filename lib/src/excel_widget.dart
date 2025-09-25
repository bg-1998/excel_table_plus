import 'dart:math';

import 'package:flutter/material.dart';
import 'package:excel_table_plus/src/ext/ext.dart';
import 'package:excel_table_plus/src/widget/excel_line.dart';
import 'package:excel_table_plus/src/excel_controller.dart';
import 'excel_model.dart';

class ExcelWidget extends StatefulWidget {

  final Widget? Function(int x, int y, ExcelItemModel? model)? itemBuilder;
  final ExcelController controller;

  /// [itemBuilder] 自定义单元格构建器
  /// [controller] Excel控制器
  const ExcelWidget({
    super.key,
    this.itemBuilder,
    required this.controller,
  });

  @override
  State<ExcelWidget> createState() => _ExcelWidgetState();
}

class _ExcelWidgetState extends State<ExcelWidget> {
  double snW = 0;
  double snH = 0;

  @override
  void initState() {
    super.initState();
    _onScrollListener();
    // 监听控制器变化
    widget.controller.addListener(_onControllerChanged);
  }
  
  @override
  void dispose() {
    // 移除监听器
    widget.controller.removeListener(_onControllerChanged);
    if(widget.controller.isAutoDispose){
      widget.controller.dispose();
    } else {
      widget.controller.stopAutoScrollTimer();
    }
    super.dispose();
  }
  
  void _onControllerChanged() {
    // 控制器状态变化时更新UI
    if (mounted) {
      setState(() {
        // 状态已在控制器中更新，这里只需要触发重建
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    final excel = widget.controller.excel;
    final items = widget.controller.items;

    // 数据校验
    if (excel.x <= 0 || excel.y <= 0) {
      throw Exception('Excel dimensions must be greater than zero. Current dimensions: ${excel.x} x ${excel.y}');
    }

    if (excel.itemWidth <= 0 || excel.itemHeight <= 0) {
      throw Exception('Excel item dimensions must be greater than zero. Current item dimensions: ${excel.itemWidth} x ${excel.itemHeight}');
    }

    if (excel.customColumnWidths.length != excel.x) {
      throw Exception('Custom column widths list length (${excel.customColumnWidths.length}) does not match the number of columns (${excel.x})');
    }

    if (excel.customRowHeights.length != excel.y) {
      throw Exception('Custom row heights list length (${excel.customRowHeights.length}) does not match the number of rows (${excel.y})');
    }

    for (int i = 0; i < excel.customColumnWidths.length; i++) {
      if (excel.customColumnWidths[i] <= 0) {
        throw Exception('Column $i width must be greater than zero. Current value: ${excel.customColumnWidths[i]}');
      }
    }

    for (int i = 0; i < excel.customRowHeights.length; i++) {
      if (excel.customRowHeights[i] <= 0) {
        throw Exception('Row $i height must be greater than zero. Current value: ${excel.customRowHeights[i]}');
      }
    }

    for (final item in items) {
      if (item.position.x < 0 || item.position.x >= excel.x) {
        throw Exception('Item position x (${item.position.x}) is out of bounds. Excel has ${excel.x} columns');
      }

      if (item.position.y < 0 || item.position.y >= excel.y) {
        throw Exception('Item position y (${item.position.y}) is out of bounds. Excel has ${excel.y} rows');
      }

      if (item.isMergeCell) {
        if (item.positions.isEmpty) {
          throw Exception('Merge cell at position (${item.position.x}, ${item.position.y}) must have positions');
        }

        for (final pos in item.positions) {
          if (pos.x < 0 || pos.x >= excel.x) {
            throw Exception('Merge cell position x ($pos.x) is out of bounds. Excel has ${excel.x} columns');
          }

          if (pos.y < 0 || pos.y >= excel.y) {
            throw Exception('Merge cell position y ($pos.y) is out of bounds. Excel has ${excel.y} rows');
          }
        }
      }
    }

    return LayoutBuilder(
      builder: (context, constraints) {
        final screenSize = Size(constraints.maxWidth, constraints.maxHeight);
        snW = (excel.sn?.itemWidth ?? excel.itemWidth) + excel.dividerWidth;
        snH = (excel.sn?.itemHeight ?? excel.itemHeight) + excel.dividerWidth;
        final constraintWidth = screenSize.width;
        final constraintHeight = screenSize.height;
        widget.controller.excelWidth = widget.controller.getExcelWidth();
        widget.controller.excelHeight = widget.controller.getExcelHeight();
        double totalExcelWidth = widget.controller.excelWidth;
        double totalExcelHeight = widget.controller.excelHeight;
        double width = totalExcelWidth;
        double height = totalExcelHeight;
        if(excel.showSn){
          totalExcelWidth += snW;
          totalExcelHeight += snH;
          if(width>constraintWidth){
            width += snW;
          }
          if(height>constraintHeight){
            height += snH;
          }
        } else {
          snW = 0;
          snH = 0;
        }
        return SizedBox(
          width: min(totalExcelWidth, constraintWidth),
          height: min(totalExcelHeight, constraintHeight),
          child: Stack(
            clipBehavior: Clip.none,
            children: [
              ClipRRect(
                borderRadius: BorderRadius.circular(excel.borderRadius),
                child: ScrollbarTheme(
                  data: const ScrollbarThemeData(
                    thickness: WidgetStatePropertyAll(0.0),
                  ),
                  child: SizedBox(
                    width: min(totalExcelWidth, constraintWidth),
                    height: min(totalExcelHeight, constraintHeight),
                    child: Stack(
                      children: [
                        Positioned(
                          left: snW - excel.dividerWidth,
                          top: 0,
                          child: SizedBox(
                            width: min(totalExcelWidth, constraintWidth) - (snW - excel.dividerWidth),
                            height: snH,
                            child: ListView.builder(
                              scrollDirection: Axis.horizontal,
                              controller: widget.controller.snHorizontalController,
                              padding: EdgeInsets.zero,
                              physics: const NeverScrollableScrollPhysics(),
                              itemCount: excel.showSn ? excel.x : 0,
                              itemBuilder: _buildHorizontalSnLineItems,
                            ),
                          ),
                        ),
                        Positioned(
                          top: snH - excel.dividerWidth,
                          left: 0,
                          child: SizedBox(
                            width: snW,
                            height: min(totalExcelHeight, constraintHeight) - (snH - excel.dividerWidth),
                            child: ListView.builder(
                              controller: widget.controller.snVerticalController,
                              padding: EdgeInsets.zero,
                              physics: const NeverScrollableScrollPhysics(),
                              itemCount: excel.showSn ? excel.y : 0,
                              itemBuilder: _buildVerticalSnLineItems,
                            ),
                          ),
                        ),
                        Positioned(
                          left: excel.showSn ? snW : 0,
                          top: excel.showSn ? snH: 0,
                          child: SizedBox(
                            height: min(widget.controller.excelHeight, constraintHeight),
                            child: SingleChildScrollView(
                              physics: const ClampingScrollPhysics(),
                              controller: widget.controller.excelVerticalController,
                              child: SizedBox(
                                width: min(widget.controller.excelWidth, constraintWidth),
                                child: SingleChildScrollView(
                                  physics: const ClampingScrollPhysics(),
                                  controller: widget.controller.excelHorizontalController,
                                  scrollDirection: Axis.horizontal,
                                  child: GestureDetector(
                                    onPanStart: widget.controller.selectedPosition!=null?widget.controller.onPanStart:null,
                                    onPanUpdate: widget.controller.selectedPosition!=null?(details) {
                                      final areaSize = Size(min(widget.controller.excelWidth, constraintWidth), min(widget.controller.excelHeight, constraintHeight));
                                      widget.controller.onPanUpdate(details,areaSize);
                                    }:null,
                                    onPanEnd: widget.controller.selectedPosition!=null?widget.controller.onPanEnd:null,
                                    child: Stack(
                                      children: [
                                        Container(
                                          width: width,
                                          height: height,
                                          decoration: BoxDecoration(
                                            color: excel.backgroundColor,
                                          ),
                                          child: Stack(
                                            children: _buildExcelLinesCells(excel, items, width, height),
                                          ),
                                        ),
                                        // 添加框选辅助框
                                        if (widget.controller.selectionRect != null)
                                          Listener(
                                            onPointerDown: (_) {
                                              widget.controller.clearMultipleSelected();
                                            },
                                            behavior: HitTestBehavior.translucent,
                                            child: IgnorePointer(
                                              child: SizedBox(
                                                width: width,
                                                height: height,
                                                child: Stack(
                                                  children: [
                                                    Positioned(
                                                      left: widget.controller.selectionRect!.left,
                                                      top: widget.controller.selectionRect!.top,
                                                      child: Container(
                                                        width: widget.controller.selectionRect!.width,
                                                        height: widget.controller.selectionRect!.height,
                                                        decoration: BoxDecoration(
                                                          border: Border.all(
                                                              color: excel.selectedBorderColor??Theme.of(context).primaryColor.withValues(alpha: 0.8),
                                                              width: excel.selectedBorderWidth,
                                                              strokeAlign: BorderSide.strokeAlignInside
                                                          ),
                                                          color: (excel.selectedBorderColor??Theme.of(context).primaryColor).withValues(alpha: 0.1),
                                                        ),
                                                      ),
                                                    ),
                                                  ],
                                                ),
                                              ),
                                            ),
                                          ),
                                      ],
                                    ),
                                  ),
                                ),
                              ),
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),
                ),
              ),
              IgnorePointer(
                child: Container(
                  width: min(totalExcelWidth, constraintWidth),
                  height: min(totalExcelHeight, constraintHeight),
                  decoration: BoxDecoration(
                    borderRadius: BorderRadius.circular(excel.borderRadius),
                    border: Border.all(
                      color: (excel.borderColor ?? excel.dividerColor)??Colors.transparent,
                      width: excel.borderWidth,
                      strokeAlign: BorderSide.strokeAlignOutside,
                    ),
                  ),
                ),
              ),
            ],
          ),
        );
      },
    );
  }

  ///
  /// build excel lines and cells
  /// [excel] excel model
  /// [items] excel items
  /// [totalWidth] total width
  /// [totalHeight] total height
  List<Widget> _buildExcelLinesCells(ExcelModel excel, List<ExcelItemModel> items, double totalWidth, double totalHeight) {
    List<Widget> widgets = <Widget>[];
    double left = 0;
    double top = 0;
    double itemWidth = excel.itemWidth;
    double itemHeight = excel.itemHeight;

    List<Map> mergeItems = [];
    for (int i = 0; i < excel.x; i++) {
      itemWidth = widget.controller.getColumnWidth(i);
      top = 0;
      for (int j = 0; j < excel.y; j++) {
        itemHeight = widget.controller.getRowHeight(j);
        var model = items.flutterExcelFirstWhereOrNull(
                (e) => e.position.x == i && e.position.y == j);
        if(model==null||!model.isMergeCell){
          Widget? widgetItem = _itemBuilder(excel, items, i, j, left, top, model: model);
          if (widgetItem != null) {
            widgets.add(widgetItem);
          }
        } else if(model.isMergeCell){
          mergeItems.add({'model': model,'x': i, 'y': j, 'left': left, 'top': top});
        }
        top += (itemHeight + excel.dividerWidth);
      }
      left += (itemWidth + excel.dividerWidth);
    }

    left = -excel.dividerWidth;
    for (int i = 0; i <= excel.x; i++) {
      itemWidth = widget.controller.getColumnWidth(i);
      // 划线
      if (excel.dividerWidth > 0) {
        final verticalLine = Positioned(
          left: left -
              (widget.controller.isDragging ? (itemWidth - excel.dividerWidth) / 2 : 0),
          top: 0.0,
          child: Container(
            width: widget.controller.isDragging ? itemWidth : excel.dividerWidth,
            alignment: Alignment.center,
            child: IgnorePointer(
              ignoring: widget.controller.isVerticalDragging,
              child: ExcelLine(
                axis: ExcelLineAxis.vertical,
                thickness: excel.dividerWidth,
                color: Colors.transparent,
                highlightColor: excel.selectedBorderColor??Theme.of(context).primaryColor,
                length: totalHeight,
                resizable: excel.resizable && i > 0,
                index: i,
                onResize: excel.resizable && i > 0
                    ? (delta) => widget.controller.onColumnResize(i - 1, delta)
                    : null,
                onDragStateChanged: widget.controller.onDragStateChanged,
              ),
            ),
          ),
        );
        widgets.add(verticalLine);
      }
      left += (itemWidth + excel.dividerWidth);
    }

    top = -excel.dividerWidth;
    if (excel.dividerWidth > 0) {
      // 划线
      for (int i = 0; i <= excel.y; i++) {
        itemHeight = widget.controller.getRowHeight(i);
        final horizontalLine = Positioned(
          left: 0.0,
          top: top -
              (widget.controller.isDragging ? (itemHeight - excel.dividerWidth) / 2 : 0),
          child: Container(
            height: widget.controller.isDragging ? itemHeight : excel.dividerWidth,
            alignment: Alignment.center,
            child: IgnorePointer(
              ignoring: widget.controller.isHorizontalDragging,
              child: ExcelLine(
                thickness: excel.dividerWidth,
                color: Colors.transparent,
                highlightColor: excel.selectedBorderColor??Theme.of(context).primaryColor,
                length: totalWidth,
                axis: ExcelLineAxis.horizontal,
                resizable: excel.resizable && i > 0,
                index: i,
                onResize: excel.resizable && i > 0
                    ? (delta) => widget.controller.onRowResize(i - 1, delta)
                    : null,
                onDragStateChanged: widget.controller.onDragStateChanged,
              ),
            ),
          ),
        );
        widgets.add(horizontalLine);
        top += (itemHeight + excel.dividerWidth);
      }
    }

    for (var e in mergeItems) {
      Widget? widgetItem = _itemBuilder(excel, items,
          e['x'], e['y'],
          e['left'],
          e['top'], model: e['model']);
      if (widgetItem != null) {
        widgets.add(widgetItem);
      }
    }
    return widgets;
  }

  ///
  /// build excel cell item
  /// [excel] excel model
  /// [items] excel items
  /// [x] x position
  /// [y] y position
  /// [left] left position
  /// [top] top position
  /// [model] excel item model
  Widget? _itemBuilder(ExcelModel excel, List<ExcelItemModel> items, int x, int y, double left, double top,
      {ExcelItemModel? model}) {
    Color? color;
    if (excel.rowColor != null) {
      color = excel.rowColor;
    } else if (excel.columnColor != null) {
      color = excel.columnColor;
    }
    double width = widget.controller.getColumnWidth(x);
    double height = widget.controller.getRowHeight(y);
    if (model != null) {
      if (model.isMergeCell) {
        if (model.positions.isNotEmpty) {
          if (!(x == model.position.x && y == model.position.y)) {
            return null;
          }
          List<int> xs = _getMergeMinMaxX(model.positions);
          List<int> ys = _getMergeMinMaxY(model.positions);
          width = 0;
          height = 0;
          for (int i = xs.first; i <= xs.last; i++) {
            double itemWidth = widget.controller.getColumnWidth(i);
            width += (itemWidth + excel.dividerWidth);
          }
          width -= excel.dividerWidth;
          for (int i = ys.first; i <= ys.last; i++) {
            double itemHeight = widget.controller.getRowHeight(i);
            height += (itemHeight + excel.dividerWidth);
          }
          height -= excel.dividerWidth;
          color = model.color;
        }
      }
    } else {
      List<ExcelItemModel> merges =
          items.where((e) => e.isMergeCell).toList();
      for (var item in merges) {
        if (item.positions.where((e) => e.x == x && e.y == y).isNotEmpty) {
          return null;
        }
      }
    }
    ExcelPosition position = ExcelPosition(x, y);
    bool isSelectsContains = widget.controller.selectedItems.contains(position);
    return Positioned(
      left: left,
      top: top,
      child: IgnorePointer(
        ignoring: widget.controller.isDragging ? true : false,
        child: MouseRegion(
          cursor: SystemMouseCursors.click,
          child: GestureDetector(
            behavior: HitTestBehavior.translucent,
            onTap: () {
              widget.controller.clearMultipleSelected();
              widget.controller.selectPosition(position);
            },
            child: Stack(
              clipBehavior: Clip.none,
              children: [
                Container(
                  width: width,
                  height: height,
                  color: model?.color ?? color,
                  alignment: excel.alignment,
                  child: IgnorePointer(
                      ignoring: (widget.controller.selectedPosition != position||(model?.isReadOnly??false)),
                      child: _itemContentBuilder(excel, x, y, model)),
                ),
                if(x<excel.x-1)
                  Positioned(right: -excel.dividerWidth,child: IgnorePointer(
                    child: Container(
                      width: excel.dividerWidth,
                      height: height + (y<excel.y-1?excel.dividerWidth:0),
                      color: excel.dividerColor??Colors.transparent,
                    ),
                  )),
                if(y<excel.y-1)
                  Positioned(bottom: -excel.dividerWidth,child: IgnorePointer(
                    child: Container(
                      width: width+(x<excel.x-1?excel.dividerWidth:0),
                      height: excel.dividerWidth,
                      color: excel.dividerColor??Colors.transparent,
                    ),
                  )),
                IgnorePointer(
                  child: Container(
                    width: width,
                    height: height,
                    decoration: (widget.controller.selectedPosition == position ||
                        (isSelectsContains&&widget.controller.isMultiSelecting))?BoxDecoration(
                      border: Border.all(
                        color: excel.selectedBorderColor??Theme.of(context).primaryColor.withValues(alpha: 0.8),
                        width: excel.selectedBorderWidth,
                        strokeAlign: BorderSide.strokeAlignInside,
                      ),
                    ):null,
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }

  ///
  /// build excel cell item content
  /// [excel] excel model
  /// [x] x position
  /// [y] y position
  /// [model] excel item model
  Widget _itemContentBuilder(ExcelModel excel, int x, int y, ExcelItemModel? model) {
    if (widget.itemBuilder != null) {
      final widgetItem = widget.itemBuilder!(x, y, model);
      if (widgetItem != null) {
        return widgetItem;
      }
    }
    return const SizedBox();
  }

  ///
  /// get merge min max x
  /// [position] merge positions
  /// return [minX,maxX]
  List<int> _getMergeMinMaxX(List<ExcelPosition> position) {
    position.sort((a, b) => a.x.compareTo(b.x));
    int minX = position.first.x;
    int maxX = position.last.x;
    return [minX, maxX];
  }

  ///
  /// get merge min max y
  /// [position] merge positions
  /// return [minY,maxY]
  List<int> _getMergeMinMaxY(List<ExcelPosition> position) {
    position.sort((a, b) => a.y.compareTo(b.y));
    int minY = position.first.y;
    int maxY = position.last.y;
    return [minY, maxY];
  }

}

extension FlutterExcelSnWidget on _ExcelWidgetState {
  /// 纵向 1,2,3...
  Widget _buildVerticalSnLineItems(BuildContext context, int index) {
    final excel = widget.controller.excel;
    if (!excel.showSn) {
      return const SizedBox.shrink();
    }

    final dividerColor =
        excel.sn?.dividerColor ?? excel.dividerColor;
    final itemHeight = widget.controller.getRowHeight(index);

    return Column(
      children: [
        if (index == 0)
          ExcelLine(
            thickness: excel.dividerWidth,
            color: dividerColor,
            length: snW,
          ),
        _buildSnItem(
          excel,
          index,
          width: snW,
          height: itemHeight,
          dividerColor: dividerColor,
          alignment: Alignment.center,
        ),
        if (index != (excel.y-1))
        ExcelLine(
          thickness: excel.dividerWidth,
          color: dividerColor,
          length: snW,
        ),
      ],
    );
  }

  /// 横向 A,B,C...
  Widget _buildHorizontalSnLineItems(BuildContext context, int index) {
    final excel = widget.controller.excel;
    if (!excel.showSn) {
      return const SizedBox.shrink();
    }

    final dividerColor =
        excel.sn?.dividerColor ?? excel.dividerColor;
    final itemWidth = widget.controller.getColumnWidth(index);

    return Row(
      children: [
        if (index == 0)
          ExcelLine(
            axis: ExcelLineAxis.vertical,
            thickness: excel.dividerWidth,
            color: dividerColor,
            length: snH,
          ),
        _buildSnItem(
          excel,
          index,
          height: snH,
          width: itemWidth,
          dividerColor: dividerColor,
          convert: true,
          alignment: Alignment.center,
        ),
        if (index != (excel.x-1))
          ExcelLine(
            axis: ExcelLineAxis.vertical,
            thickness: excel.dividerWidth,
            color: dividerColor,
            length: snH,
          ),
      ],
    );
  }

  ///
  /// build sn item
  Widget _buildSnItem(
    ExcelModel excel,
    int index, {
    required double width,
    required double height,
    Color? dividerColor,
    bool convert = false,
    Alignment alignment = Alignment.bottomCenter,
  }) {
    String value = '${index + 1}';
    if (convert) {
      index = index + 1;
      value = _convertSn(index);
    }
    return MouseRegion(
      cursor: SystemMouseCursors.click,
      child: GestureDetector(
        onTap: () {
          // 点击行号选择整行，点击列号选择整列
          if (convert) {
            // 点击的是列标题(A,B,C...)
            widget.controller.selectEntireColumn(index - 1);
          } else {
            // 点击的是行标题(1,2,3...)
            widget.controller.selectEntireRow(index);
          }
        },
        child: SizedBox(
          width: width,
          height: height,
          child: Stack(
            clipBehavior: Clip.none,
            children: [
              Container(
                color: excel.sn?.backgroundColor,
                width: width,
                height: height,
                alignment: alignment,
                child: FittedBox(
                  child: Text(
                    value,
                    style: excel.sn?.style,
                  ),
                ),
              ),
              if(convert)
              Positioned(bottom: 0,child: Container(
                width: width,
                height: excel.dividerWidth,
                color: dividerColor,
              )) else Positioned(right: 0,child: Container(
                width: excel.dividerWidth,
                height: height,
                color: dividerColor,
              )),
            ],
          ),
        ),
      ),
    );
  }

  ///
  /// convert sn index to A,B,C...
  String _convertSn(int index) {
    String result = '';
    while (index > 0) {
      index--;
      int temp = index % 26;
      result = String.fromCharCode(65 + temp) + result;
      index ~/= 26;
    }
    return result;
  }
}

extension FlutterExcelWidgetScroll on _ExcelWidgetState {
  ///
  /// on scroll listener
  void _onScrollListener() {
    widget.controller.excelHorizontalController.addListener(_onHorizontalScrolled);
    widget.controller.excelVerticalController.addListener(_onVerticalScrolled);
  }

  ///
  /// on horizontal scrolled set sn horizontal scroll to position
  void _onHorizontalScrolled() =>
      widget.controller.snHorizontalController.jumpTo(widget.controller.excelHorizontalController.offset);

  ///
  /// on vertical scrolled set sn vertical scroll to position
  void _onVerticalScrolled() =>
      widget.controller.snVerticalController.jumpTo(widget.controller.excelVerticalController.offset);
}