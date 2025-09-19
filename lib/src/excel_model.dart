import 'package:flutter/material.dart';

final Map<String, Alignment> _alignmentMap = {
  'Alignment.centerLeft': Alignment.centerLeft,
  'Alignment.center': Alignment.center,
  'Alignment.centerRight': Alignment.centerRight,
  'Alignment.topLeft': Alignment.topLeft,
  'Alignment.topCenter': Alignment.topCenter,
  'Alignment.topRight': Alignment.topRight,
  'Alignment.bottomLeft': Alignment.bottomLeft,
  'Alignment.bottomCenter': Alignment.bottomCenter,
  'Alignment.bottomRight': Alignment.bottomRight,
};
class ExcelModel {
  int x;
  int y;

  double itemWidth; // 列宽
  double itemHeight; // 行高
  List<double> customColumnWidths;
  List<double> customRowHeights;

  double selectedBorderWidth;// 选中单元格边框宽度
  Color? selectedBorderColor;// 选中单元格边框颜色
  double dividerWidth; // 分割线宽度
  double borderWidth; // 表格边线宽度
  double borderRadius; // 表格边线圆角
  Color? borderColor; // 表格边线颜色

  Color? rowColor; // 行颜色
  Color? columnColor; // 列颜色

  Color? backgroundColor; // 背景颜色
  Color? dividerColor; // 分割线颜色
  Color? intersectionColor; // 横竖交集的颜色

  bool isReadOnly; // 是否只读,此处为只读的时候,单元格的 isReadOnly 属性将无效
  bool isEnableMultipleSelection; // 是否启用多选

  bool showSn; // 是否显示序号
  ExcelSnModel? sn; // 序号
  Alignment alignment;
  
  // 添加拖拽调整行列宽高的属性
  bool resizable; // 是否可调整大小
  double resizeAreaSize; // 调整区域大小

  ///
  /// [x] 横轴坐标
  /// [y] 纵轴坐标
  /// [itemWidth] 列宽
  /// [itemHeight] 行高
  /// [selectedBorderWidth] 选中单元格边框宽度
  /// [selectedBorderColor] 选中单元格边框颜色
  /// [dividerWidth] 分割线宽度
  /// [borderWidth] 表格边线宽度
  /// [borderRadius] 表格边线圆角
  /// [borderColor] 表格边线颜色
  /// [rowColor] 行颜色
  /// [columnColor] 列颜色
  /// [neighborColors] 邻格颜色
  /// [backgroundColor] 背景颜色
  /// [noNeighborColorPositions] 不需要渲染邻格颜色的位置
  /// [isReadOnly] 是否只读
  /// [showSn] 是否显示序号
  /// [sn] 序号
  /// [alignment] 对齐方式
  /// [dividerColor] 分割线颜色
  /// [intersectionColor] 横竖交集的颜色
  /// [positionColor] 位置颜色条件
  /// [rowHeight] 按行条件设置高度
  /// [columnWidth] 按列条件设置宽度
  /// [resizable] 是否可调整行列宽高
  /// [resizeAreaSize] 调整区域大小（像素）
  ///
  ExcelModel({
    required this.x,
    required this.y,
    this.itemWidth = 80.0,
    this.itemHeight = 36.0,
    List<double>? customColumnWidths,
    List<double>? customRowHeights,
    this.selectedBorderWidth = 1.0,
    this.selectedBorderColor,
    this.dividerWidth = 0.5,
    this.borderWidth = 0.5,
    this.borderRadius = 0.0,
    this.borderColor,
    this.rowColor,
    this.columnColor,
    this.backgroundColor,
    Color? dividerColor,
    this.showSn = false,
    ExcelSnModel? sn,
    this.isReadOnly = false,
    this.isEnableMultipleSelection = true,
    this.intersectionColor,
    this.alignment = Alignment.centerLeft,
    this.resizable = false, // 默认不可调整
    this.resizeAreaSize = 5.0, // 默认调整区域为5像素
  })  : sn = sn ?? ExcelSnModel(),
        dividerColor = dividerColor ?? Colors.black.withValues(alpha: 0.15),
        customColumnWidths = customColumnWidths ?? List.filled(x, itemWidth),
        customRowHeights = customRowHeights ?? List.filled(y, itemHeight);

  /// 将ExcelModel转换为JSON对象
  Map<String, dynamic> toJson() {
    final Map<String, dynamic> json = <String, dynamic>{};
    json['x'] = x;
    json['y'] = y;
    json['itemWidth'] = itemWidth;
    json['itemHeight'] = itemHeight;
    json['customColumnWidths'] = customColumnWidths;
    json['customRowHeights'] = customRowHeights;
    json['selectedBorderWidth'] = selectedBorderWidth;
    json['selectedBorderColor'] = selectedBorderColor?.value;
    json['dividerWidth'] = dividerWidth;
    json['borderWidth'] = borderWidth;
    json['borderRadius'] = borderRadius;
    json['borderColor'] = borderColor?.value;
    json['rowColor'] = rowColor?.value;
    json['columnColor'] = columnColor?.value;
    json['backgroundColor'] = backgroundColor?.value;
    json['dividerColor'] = dividerColor?.value;
    json['intersectionColor'] = intersectionColor?.value;
    json['isReadOnly'] = isReadOnly;
    json['isEnableMultipleSelection'] = isEnableMultipleSelection;
    json['showSn'] = showSn;
    // 注意：sn (ExcelSnModel) 字段不包含在序列化中
    json['alignment'] = alignment.toString();
    json['resizable'] = resizable;
    json['resizeAreaSize'] = resizeAreaSize;
    return json;
  }

  /// 从JSON对象创建ExcelModel实例
  factory ExcelModel.fromJson(Map<String, dynamic> json) {
    // 解析Alignment
    String alignmentString = json['alignment'] as String;
    Alignment alignment = _alignmentMap[alignmentString] ??Alignment.centerLeft;
    return ExcelModel(
      x: json['x'] as int,
      y: json['y'] as int,
      itemWidth: json['itemWidth'] as double,
      itemHeight: json['itemHeight'] as double,
      customColumnWidths: List<double>.from(json['customColumnWidths'] as List),
      customRowHeights: List<double>.from(json['customRowHeights'] as List),
      selectedBorderWidth: json['selectedBorderWidth'] as double,
      selectedBorderColor: json['selectedBorderColor'] != null ? Color(json['selectedBorderColor'] as int) : null,
      dividerWidth: json['dividerWidth'] as double,
      borderRadius: json['borderRadius'] as double,
      borderWidth: json['borderWidth'] as double,
      borderColor: json['borderColor'] != null ? Color(json['borderColor'] as int) : null,
      rowColor: json['rowColor'] != null ? Color(json['rowColor'] as int) : null,
      columnColor: json['columnColor'] != null ? Color(json['columnColor'] as int) : null,
      backgroundColor: Color(json['backgroundColor'] as int),
      dividerColor: json['dividerColor'] != null ? Color(json['dividerColor'] as int) : null,
      intersectionColor: json['intersectionColor'] != null ? Color(json['intersectionColor'] as int) : null,
      isReadOnly: json['isReadOnly'] as bool,
      isEnableMultipleSelection: json['isEnableMultipleSelection'] as bool,
      showSn: json['showSn'] as bool,
      // 注意：sn (ExcelSnModel) 字段不包含在反序列化中
      alignment: alignment,
      resizable: json['resizable'] as bool,
      resizeAreaSize: json['resizeAreaSize'] as double,
    );
  }
}

/// 序号属性,参考ExcelModel
class ExcelSnModel {
  final double itemWidth;
  final double itemHeight;
  final Color? backgroundColor;
  final Color? dividerColor;
  final TextStyle? style;

  ExcelSnModel({
    this.itemWidth = 20.0,
    this.itemHeight = 20.0,
    Color? backgroundColor,
    Color? dividerColor,
    TextStyle? style,
  })  : backgroundColor = backgroundColor ?? Colors.black.withValues(alpha: 0.02),
        dividerColor = dividerColor ?? Colors.black.withValues(alpha: 0.15),
        style = style ??
            TextStyle(
                fontSize: 14.0,
                color: const Color(0xFF000000).withValues(alpha: .45));
}

class ExcelItemModel{
  ExcelPosition position; // 在作为合并单元格的时候，position代表左上角的位置
  List<ExcelPosition> positions;
  bool isMergeCell; // 是否是合并单元格,如果是合并单元格,则 positions 不能为空
  String? cellType; // 单元格类型，用于标识用户自定义的单元格model类
  dynamic value;
  bool isReadOnly;
  Color? color;

  ExcelItemModel({
    required this.position,
    this.value,
    this.positions = const <ExcelPosition>[],
    this.isMergeCell = false,
    this.cellType, // 添加cellType字段
    this.isReadOnly = false,
    this.color,
  });

  ExcelItemModel copyWith({
    ExcelPosition? position,
    dynamic value,
    List<ExcelPosition>? positions,
    bool? isMergeCell,
    String? cellType,
    bool? isReadOnly,
    Color? color,
  }) {
    return ExcelItemModel(
      position: position ?? ExcelPosition(this.position.x, this.position.y),
      value: value ?? this.value,
      positions: positions != null 
        ? positions.map((p) => ExcelPosition(p.x, p.y)).toList() 
        : this.positions.map((p) => ExcelPosition(p.x, p.y)).toList(),
      isMergeCell: isMergeCell ?? this.isMergeCell,
      cellType: cellType ?? this.cellType,
      isReadOnly: isReadOnly ?? this.isReadOnly,
      color: color ?? this.color,
    );
  }

  Map<String, dynamic> toJson() {
    final Map<String, dynamic> json = <String, dynamic>{};
    json['position'] = position.toJson();
    json['positions'] = positions.map((e) => e.toJson()).toList();
    json['isMergeCell'] = isMergeCell;
    json['cellType'] = cellType; // 添加cellType到JSON
    
    // 处理 value 字段的序列化
    if (value is Map<String, dynamic>) {
      json['value'] = value;
    } else if (value is List) {
      json['value'] = value;
    } else if (value != null && value is! num && value is! bool && value is! String) {
      // 如果 value 是自定义对象且有 toJson 方法
      if (value.toJson != null) {
        json['value'] = value.toJson();
      } else {
        // 对于没有 toJson 方法的对象，转换为字符串
        json['value'] = value.toString();
      }
    } else {
      // 基本类型直接赋值
      json['value'] = value;
    }
    
    json['isReadOnly'] = isReadOnly;
    json['color'] = color?.value;
    return json;
  }
  
  /// 从JSON对象创建ExcelItemModel实例
  factory ExcelItemModel.fromJson(Map<String, dynamic> json, {Function(String cellType, Map<String, dynamic> valueJson)? customValueFactory}) {
    dynamic value;
    final cellType = json['cellType'] as String?;
    
    // 根据cellType解析value
    if (cellType != null && customValueFactory != null) {
      // 使用自定义工厂函数创建value对象
      if (json['value'] is Map<String, dynamic>) {
        value = customValueFactory(cellType, json['value'] as Map<String, dynamic>);
      } else {
        value = json['value'];
      }
    } else {
      // 默认处理
      value = json['value'];
    }

    return ExcelItemModel(
      position: ExcelPosition.fromJson(json['position'] as Map<String, dynamic>),
      value: value,
      positions: (json['positions'] as List<dynamic>)
          .map((e) => ExcelPosition.fromJson(e as Map<String, dynamic>))
          .toList(),
      isMergeCell: json['isMergeCell'] as bool,
      cellType: cellType, // 从JSON中读取cellType
      isReadOnly: json['isReadOnly'] as bool,
      color: json['color'] != null ? Color(json['color'] as int) : null,
    );
  }
}

class ExcelPosition {
  int x;
  int y;

  ExcelPosition(this.x, this.y);

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;
    if (other.runtimeType != runtimeType) return false;
    return other is ExcelPosition && other.x == x && other.y == y;
  }

  @override
  int get hashCode => Object.hash(x, y);

  Map<String, dynamic> toJson() {
    final Map<String, dynamic> json = <String, dynamic>{};
    json['x'] = x;
    json['y'] = y;
    return json;
  }
  
  /// 从JSON对象创建ExcelPosition实例
  factory ExcelPosition.fromJson(Map<String, dynamic> json) {
    return ExcelPosition(
      json['x'] as int,
      json['y'] as int,
    );
  }
}

// 添加copyWith方法到ExcelModel类
extension ExcelModelCopyWith on ExcelModel {
  ExcelModel copyWith({
    int? x,
    int? y,
    double? itemWidth,
    double? itemHeight,
    List<double>? customColumnWidths,
    List<double>? customRowHeights,
    double? selectedBorderWidth,
    Color? selectedBorderColor,
    double? dividerWidth,
    double? borderWidth,
    double? borderRadius,
    Color? borderColor,
    Color? rowColor,
    Color? columnColor,
    Color? backgroundColor,
    Color? dividerColor,
    bool? showSn,
    ExcelSnModel? sn,
    bool? isReadOnly,
    bool? isEnableMultipleSelection,
    Color? intersectionColor,
    Alignment? alignment,
    bool? resizable,
    double? resizeAreaSize,
  }) {
    return ExcelModel(
      x: x ?? this.x,
      y: y ?? this.y,
      itemWidth: itemWidth ?? this.itemWidth,
      itemHeight: itemHeight ?? this.itemHeight,
      customColumnWidths: customColumnWidths ?? List.filled(x ?? this.x, itemWidth ?? this.itemWidth),
      customRowHeights: customRowHeights ?? List.filled(y ?? this.y, itemHeight ?? this.itemHeight),
      selectedBorderWidth: selectedBorderWidth ?? this.selectedBorderWidth,
      selectedBorderColor: selectedBorderColor ?? this.selectedBorderColor,
      dividerWidth: dividerWidth ?? this.dividerWidth,
      borderWidth: borderWidth ?? this.borderWidth,
      borderRadius: borderRadius ?? this.borderRadius,
      borderColor: borderColor ?? this.borderColor,
      rowColor: rowColor ?? this.rowColor,
      columnColor: columnColor ?? this.columnColor,
      backgroundColor: backgroundColor ?? this.backgroundColor,
      dividerColor: dividerColor ?? this.dividerColor,
      showSn: showSn ?? this.showSn,
      sn: sn ?? this.sn,
      isReadOnly: isReadOnly ?? this.isReadOnly,
      isEnableMultipleSelection: isEnableMultipleSelection ?? this.isEnableMultipleSelection,
      intersectionColor: intersectionColor ?? this.intersectionColor,
      alignment: alignment ?? this.alignment,
      resizable: resizable ?? this.resizable,
      resizeAreaSize: resizeAreaSize ?? this.resizeAreaSize,
    );
  }
}