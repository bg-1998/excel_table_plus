import 'package:flutter/material.dart';

class ResizeHandlePainter extends CustomPainter {
  @override
  void paint(Canvas canvas, Size size) {
    final paint = Paint()
      ..color = Colors.grey
      ..strokeWidth = 2
      ..style = PaintingStyle.stroke;

    final double padding = 4;
    final double handleSize = 8;

    // 绘制对角线手柄
    canvas.drawLine(
      Offset(padding, size.height - padding),
      Offset(size.width - padding, padding),
      paint,
    );

    // 绘制辅助线
    canvas.drawLine(
      Offset(padding + handleSize, size.height - padding),
      Offset(size.width - padding, padding + handleSize),
      paint,
    );

    canvas.drawLine(
      Offset(padding, size.height - padding - handleSize),
      Offset(size.width - padding - handleSize, padding),
      paint,
    );

    canvas.drawLine(
      Offset(padding + handleSize, size.height - padding),
      Offset(size.width - padding, padding + handleSize),
      paint,
    );

    canvas.drawLine(
      Offset(padding, size.height - padding - handleSize),
      Offset(size.width - padding - handleSize, padding),
      paint,
    );
  }

  @override
  bool shouldRepaint(covariant CustomPainter oldDelegate) => false;
}