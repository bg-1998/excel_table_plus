import 'dart:math';

import 'package:flutter/material.dart';

enum ExcelLineAxis {
  horizontal,
  vertical,
}

// 添加拖拽状态枚举
enum DragState {
  started,
  ended
}

class ExcelLine extends StatelessWidget {
  final Color? color;
  final double prefix;
  final double suffix;
  final ExcelLineAxis axis;
  final double thickness;
  final double length;
  
  // 添加拖拽调整大小相关的属性
  final bool resizable;
  final Function(double delta)? onResize;
  // 添加拖拽状态回调
  final Function(DragState state,ExcelLineAxis axis)? onDragStateChanged;
  final int index;

  ///
  /// [color] divider color
  /// [prefix] prefix spacing
  /// [suffix] suffix spacing
  /// [axis] axis
  /// [thickness] divider thickness
  /// [length] divider length
  /// [resizable] 是否可调整大小
  /// [onResize] 调整大小回调
  /// [onDragStateChanged] 拖拽状态改变回调
  /// [index] 索引
  ExcelLine({
    Key? key,
    Color? color,
    this.prefix = 0.0,
    this.suffix = 0.0,
    this.axis = ExcelLineAxis.horizontal,
    this.thickness = 1.0,
    this.length = double.infinity,
    this.resizable = false,
    this.onResize,
    this.onDragStateChanged,
    this.index = 0,
  })  : color = color ?? const Color(0xFFEFEFEF).withValues(alpha: 0.6),
        super(key: key);

  @override
  Widget build(BuildContext context) {
    Widget line;
    switch (axis) {
      case ExcelLineAxis.horizontal:
        line = _buildHorizontalLine();
        break;
      case ExcelLineAxis.vertical:
        line = _buildVerticalLine();
        break;
    }
    
    // 如果支持调整大小，则包装一个可以拖拽的区域
    if (resizable) {
      return _buildResizableLine(context, line);
    }
    
    return line;
  }

  Widget _buildResizableLine(BuildContext context, Widget line) {
    return MouseRegion(
      cursor: axis == ExcelLineAxis.vertical
        ? SystemMouseCursors.resizeColumn
        : SystemMouseCursors.resizeRow,
      child: GestureDetector(
        behavior: HitTestBehavior.translucent,
        onPanStart: (details) {
          // 开始拖拽
          if (onDragStateChanged != null) {
            onDragStateChanged!(DragState.started, axis);
          }
        },
        onPanUpdate: (details) {
          // 更新拖拽
          if (onResize != null) {
            double delta = axis == ExcelLineAxis.vertical
              ? details.delta.dx
              : details.delta.dy;
            onResize!(delta);
          }
        },
        onPanEnd: (details) {
          // 结束拖拽
          if (onDragStateChanged != null) {
            onDragStateChanged!(DragState.ended, axis);
          }
        },
        child: Container(
          alignment: Alignment.center,
          color: Colors.transparent,
          child: line,
        ),
      ),
    );
  }

  Widget _buildVerticalLine() {
    double height = length;
    height = height - prefix - suffix;
    height = max(0.0, height);
    return SizedBox(
      width: thickness,
      height: height,
      child: VerticalDivider(
        color: color,
        thickness: thickness,
        width: thickness,
        indent: prefix,
        endIndent: suffix,
      ),
    );
  }

  Widget _buildHorizontalLine() {
    double width = length;
    width = width - prefix - suffix;
    width = max(0.0, width);
    return SizedBox(
      width: width,
      height: thickness,
      child: Divider(
        thickness: thickness,
        height: thickness,
        color: color,
        indent: prefix,
        endIndent: suffix,
      ),
    );
  }
}
