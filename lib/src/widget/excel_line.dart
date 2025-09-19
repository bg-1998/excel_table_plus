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

class ExcelLine extends StatefulWidget {
  final Color? color;
  final Color? highlightColor;
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
  /// [highlightColor] divider highlight color
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
    super.key,
    Color? color,
    this.highlightColor,
    this.prefix = 0.0,
    this.suffix = 0.0,
    this.axis = ExcelLineAxis.horizontal,
    this.thickness = 1.0,
    this.length = double.infinity,
    this.resizable = false,
    this.onResize,
    this.onDragStateChanged,
    this.index = 0,
  })  : color = color ?? const Color(0xFFEFEFEF).withValues(alpha: 0.6);

  @override
  State<ExcelLine> createState() => _ExcelLineState();
}

class _ExcelLineState extends State<ExcelLine> {
  // 添加鼠标悬浮状态
  bool _onHover = false;

  @override
  Widget build(BuildContext context) {
    Widget line;
    switch (widget.axis) {
      case ExcelLineAxis.horizontal:
        line = _buildHorizontalLine();
        break;
      case ExcelLineAxis.vertical:
        line = _buildVerticalLine();
        break;
    }

    // 如果支持调整大小，则包装一个可以拖拽的区域
    if (widget.resizable) {
      return _buildResizableLine(context, line);
    }

    return line;
  }

  Widget _buildResizableLine(BuildContext context, Widget line) {
    return MouseRegion(
      cursor: widget.axis == ExcelLineAxis.vertical
        ? SystemMouseCursors.resizeColumn
        : SystemMouseCursors.resizeRow,
      onHover: (event) {
        setState(() {
          _onHover = true;
        });
      },
      onExit: (event) {
        setState(() {
          _onHover = false;
        });
      },
      child: GestureDetector(
        behavior: HitTestBehavior.translucent,
        onPanStart: (details) {
          setState(() {
            _onHover = true;
          });
          // 开始拖拽
          if (widget.onDragStateChanged != null) {
            widget.onDragStateChanged!(DragState.started, widget.axis);
          }
        },
        onPanUpdate: (details) {
          // 更新拖拽
          if (widget.onResize != null) {
            double delta = widget.axis == ExcelLineAxis.vertical
              ? details.delta.dx
              : details.delta.dy;
            widget.onResize!(delta);
          }
        },
        onPanEnd: (details) {
          setState(() {
            _onHover = false;
          });
          // 结束拖拽
          if (widget.onDragStateChanged != null) {
            widget.onDragStateChanged!(DragState.ended, widget.axis);
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
    double height = widget.length;
    height = height - widget.prefix - widget.suffix;
    height = max(0.0, height);
    return SizedBox(
      width: widget.thickness,
      height: height,
      child: VerticalDivider(
        color: _onHover? widget.highlightColor??widget.color:widget.color,
        thickness: widget.thickness,
        width: widget.thickness,
        indent: widget.prefix,
        endIndent: widget.suffix,
      ),
    );
  }

  Widget _buildHorizontalLine() {
    double width = widget.length;
    width = width - widget.prefix - widget.suffix;
    width = max(0.0, width);
    return SizedBox(
      width: width,
      height: widget.thickness,
      child: Divider(
        thickness: widget.thickness,
        height: widget.thickness,
        color: _onHover? widget.highlightColor??widget.color:widget.color,
        indent: widget.prefix,
        endIndent: widget.suffix,
      ),
    );
  }
}
