import 'package:flutter/material.dart';

/// 用于绘制Excel单元格边框的自定义Painter
class ExcelCellBorderPainter extends CustomPainter {
  final Color borderColor;
  final double borderWidth;

  ExcelCellBorderPainter({
    required this.borderColor,
    required this.borderWidth,
  });

  @override
  void paint(Canvas canvas, Size size) {
    final Paint paint = Paint()
      ..strokeWidth = borderWidth
      ..color = borderColor
      ..style = PaintingStyle.stroke;
    if (borderWidth > 0) {
      final Path path = Path();
      path.moveTo(size.width+borderWidth/2, -0.4);
      path.lineTo(size.width+borderWidth/2, size.height+borderWidth/2);
      path.lineTo(-0.4, size.height+borderWidth/2);
      canvas.drawPath(path, paint);
    }
  }

  @override
  bool shouldRepaint(covariant CustomPainter oldDelegate) {
    return true;
  }
}