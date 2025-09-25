import 'package:flutter/material.dart';
import 'package:excel_table_plus/src/excel_model.dart';
import 'dart:math' as math;
import 'dart:async';

import 'widget/excel_line.dart';

/// Excel列和行最小宽度
const double _defaultMinExcelColumnWidth = 5.0;
const double _defaultMinExcelRowHeight = 5.0;
const double _defaultAutoScrollThreshold = 20.0; // 触发自动滚动的距离阈值
const double _defaultMaxAutoScrollSpeed = 20.0; // 最大自动滚动速度
/// Excel控制器，用于管理Excel组件的状态和操作
/// 继承ChangeNotifier以支持widget刷新
class ExcelController extends ChangeNotifier {
  /// Excel模型
  ExcelModel _excel;
  ExcelModel get excel => _excel;
  ExcelModel? startScaleExcel;

  /// Excel项目列表
  List<ExcelItemModel> _items = [];
  List<ExcelItemModel> get items => _items;
  List<ExcelItemModel>? startScaleItems;

  /// Excel宽度
  double excelWidth = 0.0;

  /// Excel高度
  double excelHeight = 0.0;

  /// 是否自动销毁控制器
  bool isAutoDispose;

  ///Excel列和行最小宽度
  double? minExcelColumnWidth;
  double? minExcelRowHeight;
  ///最大自动滚动速度
  double? maxAutoScrollSpeed;
  ///触发自动滚动的距离阈值
  double? autoScrollThreshold;

  /// 选中的单元格位置
  ExcelPosition? _selectedPosition;
  ExcelPosition? get selectedPosition => _selectedPosition;

  /// 多选模式下选中的单元格列表
  final List<ExcelPosition> _selectedItems = [];
  List<ExcelPosition> get selectedItems => _selectedItems;

  ///添加多选相关状态
  Offset? _startPoint; // 多选起始点
  Offset? _lastPoint; // 上一个点

  ///添加框选辅助框相关变量
  Offset? _selectStartOffset;
  Offset? _selectEndOffset;

  /// 选择区域的矩形范围
  Rect? _selectionRect;
  Rect? get selectionRect => _selectionRect;

  /// 添加拖拽状态变量
  bool _isDragging = false;
  bool get isDragging => _isDragging;

  /// 添加水平拖拽状态变量
  bool _isHorizontalDragging = false;
  bool get isHorizontalDragging => _isHorizontalDragging;

  /// 添加垂直拖拽状态变量
  bool _isVerticalDragging = false;
  bool get isVerticalDragging => _isVerticalDragging;

  /// 添加多选状态变量
  bool _isMultiSelecting = false;
  bool get isMultiSelecting => _isMultiSelecting;
  
  Timer? _autoScrollTimer;
  Size? _areaSize;
  Offset _currentPoint = Offset.zero;

  /// 单元格点击回调
  Function(ExcelPosition? position)? onPositionSelected;

  /// 多选变化回调
  Function(List<ExcelPosition> selectedItems)? onMultiSelectionChanged;

  /// Excel尺寸变化回调
  Function(Size size)? onExcelSizeChanged;

  ScrollController snHorizontalController = ScrollController();
  ScrollController snVerticalController = ScrollController();
  ScrollController excelHorizontalController = ScrollController();
  ScrollController excelVerticalController = ScrollController();

  ExcelController({
    required ExcelModel excel,
    List<ExcelItemModel>? items,
    this.isAutoDispose = true,
    this.minExcelColumnWidth,
    this.minExcelRowHeight,
    this.maxAutoScrollSpeed,
    this.autoScrollThreshold,
  })  : _excel = excel,
        _items = items ?? [];

  /// 设置Excel模型
  void setExcel(ExcelModel excel) {
    _excel = excel;
    update();
  }

  /// 设置Excel项目列表
  void setItems(List<ExcelItemModel> items) {
    _items = List.from(items);
    update();
  }

  /// 更新Excel项目列表中的特定项
  void updateItem(ExcelItemModel item) {
    final index = _items.indexWhere((element) =>
        element.position.x == item.position.x &&
        element.position.y == item.position.y);
    if (index != -1) {
      _items[index] = item;
    } else {
      _items.add(item);
    }
    update();
  }

  /// 移除Excel项目列表中的特定项
  void removeItem(ExcelPosition position) {
    _items.removeWhere((element) =>
        element.position.x == position.x &&
        element.position.y == position.y);
    update();
  }

  ///执行刷新
  void update(){
    notifyListeners();
  }
  
  /// 释放资源
  @override
  void dispose() {
    stopAutoScrollTimer();
    onMultiSelectionChanged = null;
    onPositionSelected = null;
    onExcelSizeChanged = null;
    super.dispose();
  }
}

/// 控制器数据管理扩展
extension ExcelDataController on ExcelController {
  /// 设置单个选中位置
  void selectPosition(ExcelPosition? position) {
    if(_selectedPosition==position){
      _selectedPosition = null;
    } else {
      _selectedPosition = position;
    }
    update();
    
    // 触发位置选中回调
    onPositionSelected?.call(position);
  }

  /// 添加选中项（用于多选模式）
  void addSelectedItem(ExcelPosition position) {
    if (!_selectedItems.any((element) => element.x == position.x && element.y == position.y)) {
      _selectedItems.add(position);
      update();
    }
  }

  /// 移除选中项（用于多选模式）
  void removeSelectedItem(ExcelPosition position) {
    _selectedItems.removeWhere((element) => element.x == position.x && element.y == position.y);
    update();
  }

  /// 清除所有选中项
  void clearSelection() {
    _selectedPosition = null;
    _selectedItems.clear();
    _selectionRect = null;
    _selectStartOffset = null;
    _selectEndOffset = null;
    update();
  }

  /// 清除多选项
  void clearMultipleSelected() {
    _selectedItems.clear();
    _selectionRect = null;
    _selectStartOffset = null;
    _selectEndOffset = null;
    update();
  }

  /// 选择整行
  void selectEntireRow(int rowIndex) {
    if(!excel.isEnableMultipleSelection){
      return;
    }
    // 清除之前的选中状态
    _selectedItems.clear();
    _selectedPosition = null;
    // 添加整行的所有单元格到_selectedItems
    double left = 0.0;
    double top = 0.0;
    // 计算行的顶部位置
    for (int i = 0; i < rowIndex; i++) {
      top += getRowHeight(i) + excel.dividerWidth;
    }
    // 添加该行所有单元格到选中列表
    for (int col = 0; col < excel.x; col++) {
      var position = ExcelPosition(col, rowIndex);
      _selectedItems.add(position);
    }
    // 计算选择区域的矩形
    double rowHeight = getRowHeight(rowIndex);
    double totalWidth = getExcelWidth();

    _selectStartOffset = Offset(left, top);
    _selectEndOffset = Offset(left + totalWidth-excel.dividerWidth*2, top + rowHeight);
    _selectionRect = Rect.fromPoints(_selectStartOffset!, _selectEndOffset!);
    // 触发多选回调
    onMultiSelectionChanged?.call(_selectedItems);
    update();
  }

  /// 选择整列
  void selectEntireColumn(int columnIndex) {
    if(!excel.isEnableMultipleSelection){
      return;
    }
    // 清除之前的选中状态
    _selectedItems.clear();
    _selectedPosition = null;
    // 添加整列的所有单元格到_selectedItems
    double left = 0.0;
    double top = 0.0;
    // 计算列的左侧位置
    for (int j = 0; j < columnIndex; j++) {
      left += getColumnWidth(j) + excel.dividerWidth;
    }
    // 添加该列所有单元格到选中列表
    for (int row = 0; row < excel.y; row++) {
      var position = ExcelPosition(columnIndex, row);
      _selectedItems.add(position);
    }
    // 计算选择区域的矩形
    double columnWidth = getColumnWidth(columnIndex);
    double totalHeight = getExcelHeight();
    _selectStartOffset = Offset(left, top);
    _selectEndOffset = Offset(left + columnWidth, top + totalHeight - excel.dividerWidth*2);
    _selectionRect = Rect.fromPoints(_selectStartOffset!, _selectEndOffset!);
    // 触发多选回调
    onMultiSelectionChanged?.call(_selectedItems);
    update();
  }
}

/// 控制器尺寸管理扩展
extension ExcelSizeController on ExcelController{
  /// 获取Excel宽度
  double getExcelWidth() {
    double width = 0;
    for (int i = 0; i < excel.x; i++) {
      double itemWidth = getColumnWidth(i);
      width += (itemWidth + excel.dividerWidth);
    }
    width -= excel.dividerWidth;
    return width;
  }

  /// 获取Excel高度
  double getExcelHeight() {
    double height = 0;
    for (int i = 0; i < excel.y; i++) {
      double itemHeight = getRowHeight(i);
      height += (itemHeight + excel.dividerWidth);
    }
    height -= excel.dividerWidth;
    return height;
  }

  /// 获取列宽
  double getColumnWidth(int columnIndex) {
    if(columnIndex>-1&&columnIndex<excel.customColumnWidths.length){
      return excel.customColumnWidths[columnIndex];
    }
    return excel.itemWidth;
  }

  /// 获取行高
  double getRowHeight(int rowIndex) {
    if(rowIndex>-1&&rowIndex<excel.customRowHeights.length){
      return excel.customRowHeights[rowIndex];
    }
    return excel.itemHeight;
  }

  /// 设置列宽
  void setColumnWidth(int columnIndex, double width) {
    if(columnIndex>-1&&columnIndex<excel.customColumnWidths.length){
      excel.customColumnWidths[columnIndex] = width;
    }
    onExcelSizeChanged?.call(Size(getExcelWidth(), getExcelHeight()));
    update();
  }

  /// 设置行高
  void setRowHeight(int rowIndex, double height) {
    if(rowIndex>-1&&rowIndex<excel.customRowHeights.length){
      excel.customRowHeights[rowIndex] = height;
    }
    onExcelSizeChanged?.call(Size(getExcelWidth(), getExcelHeight()));
    update();
  }

  /// 处理列宽调整
  void onColumnResize(int columnIndex, double delta) {
    double currentWidth = getColumnWidth(columnIndex);
    double newWidth = currentWidth + delta;
    // 限制最小宽度
    if (newWidth < (minExcelColumnWidth??_defaultMinExcelColumnWidth)) newWidth = (minExcelColumnWidth??_defaultMinExcelColumnWidth);
    excel.customColumnWidths[columnIndex] = newWidth;
    onExcelSizeChanged?.call(Size(getExcelWidth(), getExcelHeight()));
    update();
  }

  /// 处理行高调整
  void onRowResize(int rowIndex, double delta) {
    double currentHeight = getRowHeight(rowIndex);
    double newHeight = currentHeight + delta;
    // 限制最小高度
    if (newHeight < (minExcelRowHeight??_defaultMinExcelRowHeight)) newHeight = (minExcelRowHeight??_defaultMinExcelRowHeight);
    excel.customRowHeights[rowIndex] = newHeight;
    onExcelSizeChanged?.call(Size(getExcelWidth(), getExcelHeight()));
    update();
  }

  /// 设置分割线宽度
  void setDividerWidth(double width) {
    excel.dividerWidth = width;
    onExcelSizeChanged?.call(Size(getExcelWidth(), getExcelHeight()));
    update();
  }

  /// 设置边框宽度
  void setBorderWidth(double width) {
    excel.borderWidth = width;
    onExcelSizeChanged?.call(Size(getExcelWidth(), getExcelHeight()));
    update();
  }
}

/// 控制器手势拖拽管理扩展
extension GestureDragController on ExcelController{
  /// 处理拖拽状态变化
  void onDragStateChanged(DragState state, ExcelLineAxis axis) {
    _isDragging = state == DragState.started;
    if (axis == ExcelLineAxis.horizontal) {
      _isHorizontalDragging = false;
      _isVerticalDragging = isDragging;
    } else {
      _isVerticalDragging = false;
      _isHorizontalDragging = isDragging;
    }
    update();
  }

  /// 处理拖拽开始
  void onPanStart(DragStartDetails details) {
    if(!excel.isEnableMultipleSelection){
      return;
    }
    _isMultiSelecting = true;
    _startPoint = details.localPosition;
    _lastPoint = _startPoint;
    _selectionRect = Rect.fromPoints(_startPoint!, _lastPoint!);
    _selectedItems.clear();
    _selectStartOffset = null;
    _selectEndOffset = null;
    _currentPoint = details.localPosition;
    update();
  }

  /// 处理拖拽更新
  void onPanUpdate(DragUpdateDetails details,Size areaSize) {
    if (!_isMultiSelecting) return;
    
    // 保存当前点和区域大小用于自动滚动
    _currentPoint = _currentPoint+details.delta;
    _areaSize = areaSize;
    
    // 启动或重启自动滚动定时器
    _startAutoScrollTimer();
  }

  /// 处理拖拽结束
  void onPanEnd(DragEndDetails details) {
    // 停止自动滚动定时器
    stopAutoScrollTimer();
    if (!_isMultiSelecting) return;
    _selectedPosition = null;
    _isMultiSelecting = false;
    if(_selectStartOffset!=null&&_selectEndOffset!=null){
      _selectionRect = Rect.fromPoints(_selectStartOffset!, _selectEndOffset!);
    } else {
      _selectionRect = null; // 清除框选区域
      _selectStartOffset = null;
      _selectEndOffset = null;
    }
    onMultiSelectionChanged?.call(_selectedItems);
    update();
  }

  /// 计算选中的单元格
  void _calculateSelectedCells() {
    if (_selectionRect == null) return;
    _selectedItems.clear();
    _selectStartOffset = null;
    double left = 0;
    for (int i = 0; i < excel.x; i++) {
      double itemWidth = getColumnWidth(i);
      double top = 0;
      for (int j = 0; j < excel.y; j++) {
        double itemHeight = getRowHeight(j);

        // 计算单元格的矩形区域（包含分割线）
        final cellRect = Rect.fromLTWH(
          left,
          top,
          itemWidth,
          itemHeight,
        );
        // 检查单元格是否与选择区域相交
        if (_selectionRect!.overlaps(cellRect)) {
          var position = ExcelPosition(i,j);
          _selectedItems.add(position);
          _selectStartOffset ??= cellRect.topLeft;
          if(_selectStartOffset!.dx>cellRect.topLeft.dx){
            _selectStartOffset = Offset(cellRect.topLeft.dx, _selectStartOffset!.dy);
          }
          if(_selectStartOffset!.dy>cellRect.topLeft.dy){
            _selectStartOffset = Offset(_selectStartOffset!.dx, cellRect.topLeft.dy);
          }
          _selectEndOffset = cellRect.bottomRight;
        }
        top += (itemHeight + excel.dividerWidth);
      }
      left += (itemWidth + excel.dividerWidth);
    }
  }
}

/// 控制器单元格合并管理扩展
extension MergeCellController on ExcelController {
  /// 合并选中的单元格
  void mergeSelectedCells() {
    if (_selectedItems.isEmpty) return;

    // 查找选中区域内的所有项
    List<ExcelItemModel> itemsToRemove = [];
    for (var item in _selectedItems) {
      var model = findItemContainingPosition(item.x, item.y);
      if (model != null) {
        itemsToRemove.add(model);
      }
    }

    // 从列表中移除找到的项
    _items.removeWhere((item) => itemsToRemove.contains(item));

    // 按照x,y排序
    _selectedItems.sort((a, b) {
      if (a.x != b.x) {
        return a.x.compareTo(b.x);
      }
      return a.y.compareTo(b.y);
    });

    if (_selectedItems.isNotEmpty) {
      var firstPosition = _selectedItems.first;
      ExcelItemModel newItem = ExcelItemModel(
        position: firstPosition,
        positions: List.from(_selectedItems),
        isMergeCell: true,
      );
      
      _items.add(newItem);
      _selectedItems.clear();
      _selectedItems.add(firstPosition);
      _selectedPosition = newItem.position;
      update();
    }
  }

  /// 拆分合并的单元格
  void splitMergedCell() {
    if(selectedPosition==null){
      return;
    }
    var model = findItemContainingPosition(selectedPosition!.x, selectedPosition!.y);
    if (model == null || !model.isMergeCell) {
      return;
    }
    _items.remove(model);
    _selectedPosition = model.position;
    update();
  }

  /// 查找包含指定位置的项
  ExcelItemModel? findItemContainingPosition(int x, int y) {
    for (var item in items) {
      if (item.isMergeCell) {
        if (item.positions.any((position) => position.x == x && position.y == y)) {
          return item;
        }
      } else {
        if (item.position.x == x && item.position.y == y) {
          return item;
        }
      }
    }
    return null;
  }
}

/// 控制器行列插入删除扩展
extension RowColumnInsertDeleteController on ExcelController {
  /// 在指定行索引位置插入新行
  void insertRow(int rowIndex) {
    // 更新Excel模型的行数
    final newExcel = excel.copyWith(y: excel.y + 1);
    
    // 更新自定义行高列表，插入新行的高度
    final newCustomRowHeights = List<double>.from(excel.customRowHeights);
    // 确保列表长度足够
    while (newCustomRowHeights.length < excel.y) {
      newCustomRowHeights.add(excel.itemHeight);
    }
    // 在指定位置插入新行高度
    final rowHeight = rowIndex < newCustomRowHeights.length 
        ? newCustomRowHeights[rowIndex] 
        : excel.itemHeight;
    newCustomRowHeights.insert(rowIndex, rowHeight);
    
    // 创建新的Excel模型
    final updatedExcel = newExcel.copyWith(customRowHeights: newCustomRowHeights);
    setExcel(updatedExcel);
    
    // 更新所有项目的位置
    final updatedItems = <ExcelItemModel>[];
    for (final item in _items) {
      if (item.isMergeCell) {
        // 处理合并单元格
        final updatedPositions = <ExcelPosition>[];
        bool needsUpdate = false;
        
        for (final pos in item.positions) {
          if (pos.y >= rowIndex) {
            updatedPositions.add(ExcelPosition(pos.x, pos.y + 1));
            needsUpdate = true;
          } else {
            updatedPositions.add(ExcelPosition(pos.x, pos.y));
          }
        }
        
        if (needsUpdate) {
          // 如果是合并单元格的起始位置需要调整
          final newPrimaryPosition = ExcelPosition(item.position.x, 
              item.position.y >= rowIndex ? item.position.y + 1 : item.position.y);
          
          updatedItems.add(
            ExcelItemModel(
              position: newPrimaryPosition,
              value: item.value,
              positions: updatedPositions,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      } else {
        // 处理普通单元格
        if (item.position.y >= rowIndex) {
          updatedItems.add(
            ExcelItemModel(
              position: ExcelPosition(item.position.x, item.position.y + 1),
              value: item.value,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      }
    }
    
    setItems(updatedItems);
  }

  /// 在指定列索引位置插入新列
  void insertColumn(int columnIndex) {
    // 更新Excel模型的列数
    final newExcel = excel.copyWith(x: excel.x + 1);
    
    // 更新自定义列宽列表，插入新列的宽度
    final newCustomColumnWidths = List<double>.from(excel.customColumnWidths);
    // 确保列表长度足够
    while (newCustomColumnWidths.length < excel.x) {
      newCustomColumnWidths.add(excel.itemWidth);
    }
    // 在指定位置插入新列宽度
    final columnWidth = columnIndex < newCustomColumnWidths.length 
        ? newCustomColumnWidths[columnIndex] 
        : excel.itemWidth;
    newCustomColumnWidths.insert(columnIndex, columnWidth);
    
    // 创建新的Excel模型
    final updatedExcel = newExcel.copyWith(customColumnWidths: newCustomColumnWidths);
    setExcel(updatedExcel);
    
    // 更新所有项目的位置
    final updatedItems = <ExcelItemModel>[];
    for (final item in _items) {
      if (item.isMergeCell) {
        // 处理合并单元格
        final updatedPositions = <ExcelPosition>[];
        bool needsUpdate = false;
        
        for (final pos in item.positions) {
          if (pos.x >= columnIndex) {
            updatedPositions.add(ExcelPosition(pos.x + 1, pos.y));
            needsUpdate = true;
          } else {
            updatedPositions.add(ExcelPosition(pos.x, pos.y));
          }
        }
        
        if (needsUpdate) {
          // 如果是合并单元格的起始位置需要调整
          final newPrimaryPosition = ExcelPosition(
              item.position.x >= columnIndex ? item.position.x + 1 : item.position.x,
              item.position.y);
          
          updatedItems.add(
            ExcelItemModel(
              position: newPrimaryPosition,
              value: item.value,
              positions: updatedPositions,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      } else {
        // 处理普通单元格
        if (item.position.x >= columnIndex) {
          updatedItems.add(
            ExcelItemModel(
              position: ExcelPosition(item.position.x + 1, item.position.y),
              value: item.value,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      }
    }
    
    setItems(updatedItems);
  }
  
  /// 删除指定行索引的行
  void deleteRow(int rowIndex) {
    // 检查行索引是否有效
    if (rowIndex < 0 || rowIndex >= excel.y) {
      return;
    }
    
    // 更新Excel模型的行数
    final newExcel = excel.copyWith(y: excel.y - 1);
    
    // 更新自定义行高列表，删除指定行的高度
    final newCustomRowHeights = List<double>.from(excel.customRowHeights);
    // 确保列表长度足够
    while (newCustomRowHeights.length < excel.y) {
      newCustomRowHeights.add(excel.itemHeight);
    }
    // 删除指定位置的行高度
    if (rowIndex < newCustomRowHeights.length) {
      newCustomRowHeights.removeAt(rowIndex);
    }
    
    // 创建新的Excel模型
    final updatedExcel = newExcel.copyWith(customRowHeights: newCustomRowHeights);
    _excel = updatedExcel;
    
    // 更新所有项目的位置
    final updatedItems = <ExcelItemModel>[];
    for (final item in _items) {
      if (item.isMergeCell) {
        // 处理合并单元格
        // 如果合并单元格包含要删除的行，则整个合并单元格都要删除
        if (item.positions.any((pos) => pos.y == rowIndex)) {
          continue; // 删除整个合并单元格
        }
        
        final updatedPositions = <ExcelPosition>[];
        bool needsUpdate = false;
        
        for (final pos in item.positions) {
          if (pos.y > rowIndex) {
            updatedPositions.add(ExcelPosition(pos.x, pos.y - 1));
            needsUpdate = true;
          } else {
            updatedPositions.add(ExcelPosition(pos.x, pos.y));
          }
        }
        
        if (needsUpdate) {
          // 如果是合并单元格的起始位置需要调整
          final newPrimaryPosition = ExcelPosition(item.position.x, 
              item.position.y > rowIndex ? item.position.y - 1 : item.position.y);
          
          updatedItems.add(
            ExcelItemModel(
              position: newPrimaryPosition,
              value: item.value,
              positions: updatedPositions,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      } else {
        // 处理普通单元格
        if (item.position.y == rowIndex) {
          continue; // 删除该单元格
        }
        
        if (item.position.y > rowIndex) {
          updatedItems.add(
            ExcelItemModel(
              position: ExcelPosition(item.position.x, item.position.y - 1),
              value: item.value,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      }
    }
    _items = updatedItems;
    clearSelection();
  }

  /// 删除指定列索引的列
  void deleteColumn(int columnIndex) {
    // 检查列索引是否有效
    if (columnIndex < 0 || columnIndex >= excel.x) {
      return;
    }
    
    // 更新Excel模型的列数
    final newExcel = excel.copyWith(x: excel.x - 1);
    
    // 更新自定义列宽列表，删除指定列的宽度
    final newCustomColumnWidths = List<double>.from(excel.customColumnWidths);
    // 确保列表长度足够
    while (newCustomColumnWidths.length < excel.x) {
      newCustomColumnWidths.add(excel.itemWidth);
    }
    // 删除指定位置的列宽度
    if (columnIndex < newCustomColumnWidths.length) {
      newCustomColumnWidths.removeAt(columnIndex);
    }
    
    // 创建新的Excel模型
    final updatedExcel = newExcel.copyWith(customColumnWidths: newCustomColumnWidths);
    _excel = updatedExcel;
    
    // 更新所有项目的位置
    final updatedItems = <ExcelItemModel>[];
    for (final item in _items) {
      if (item.isMergeCell) {
        // 处理合并单元格
        // 如果合并单元格包含要删除的列，则整个合并单元格都要删除
        if (item.positions.any((pos) => pos.x == columnIndex)) {
          continue; // 删除整个合并单元格
        }
        
        final updatedPositions = <ExcelPosition>[];
        bool needsUpdate = false;
        
        for (final pos in item.positions) {
          if (pos.x > columnIndex) {
            updatedPositions.add(ExcelPosition(pos.x - 1, pos.y));
            needsUpdate = true;
          } else {
            updatedPositions.add(ExcelPosition(pos.x, pos.y));
          }
        }
        
        if (needsUpdate) {
          // 如果是合并单元格的起始位置需要调整
          final newPrimaryPosition = ExcelPosition(
              item.position.x > columnIndex ? item.position.x - 1 : item.position.x,
              item.position.y);
          
          updatedItems.add(
            ExcelItemModel(
              position: newPrimaryPosition,
              value: item.value,
              positions: updatedPositions,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      } else {
        // 处理普通单元格
        if (item.position.x == columnIndex) {
          continue; // 删除该单元格
        }
        
        if (item.position.x > columnIndex) {
          updatedItems.add(
            ExcelItemModel(
              position: ExcelPosition(item.position.x - 1, item.position.y),
              value: item.value,
              isMergeCell: item.isMergeCell,
              isReadOnly: item.isReadOnly,
              color: item.color,
            ),
          );
        } else {
          updatedItems.add(item);
        }
      }
    }
    _items = updatedItems;
    clearSelection();
  }
}

/// 控制器自动滚动扩展
extension AutoScrollController on ExcelController {
  /// 启动自动滚动定时器
  void _startAutoScrollTimer() {
    Size areaSize = _areaSize!;

    // 计算当前滚动位置
    double currentHorizontalOffset = excelHorizontalController.offset;
    double currentVerticalOffset = excelVerticalController.offset;

    // 获取滚动范围
    double maxHorizontalOffset = excelHorizontalController.position.maxScrollExtent;
    double maxVerticalOffset = excelVerticalController.position.maxScrollExtent;

    bool scrolled = false;
    var autoScrollThreshold = this.autoScrollThreshold??_defaultAutoScrollThreshold;
    // 水平方向自动滚动
    if (_currentPoint.dx - currentHorizontalOffset < autoScrollThreshold && currentHorizontalOffset > 0) {
      scrolled = true;
    } else if (_currentPoint.dx > currentHorizontalOffset+areaSize.width-autoScrollThreshold && currentHorizontalOffset < maxHorizontalOffset) {
      scrolled = true;
    }
    // 垂直方向自动滚动
    if (_currentPoint.dy - currentVerticalOffset < autoScrollThreshold && currentVerticalOffset > 0) {
      scrolled = true;
    } else if (_currentPoint.dy > currentVerticalOffset + areaSize.height - autoScrollThreshold && currentVerticalOffset < maxVerticalOffset) {
      scrolled = true;
    }
    if(scrolled){
      // 启动新的定时器，每16毫秒执行一次（约60FPS）
      _autoScrollTimer ??= Timer.periodic(const Duration(milliseconds: 16), (_) {
        _performAutoScroll();
      });
    } else {
      stopAutoScrollTimer();
      _lastPoint = _currentPoint;
      _selectionRect = Rect.fromPoints(_startPoint!, _lastPoint!);
      _calculateSelectedCells();
      update();
    }
  }
  
  /// 停止自动滚动定时器
  void stopAutoScrollTimer() {
    _autoScrollTimer?.cancel();
    _autoScrollTimer = null;
  }
  
  /// 执行自动滚动
  void _performAutoScroll() {
    if (_areaSize == null) return;
    
    Size areaSize = _areaSize!;

    // 计算当前滚动位置
    double currentHorizontalOffset = excelHorizontalController.offset;
    double currentVerticalOffset = excelVerticalController.offset;

    // 获取滚动范围
    double maxHorizontalOffset = excelHorizontalController.position.maxScrollExtent;
    double maxVerticalOffset = excelVerticalController.position.maxScrollExtent;

    bool scrolled = false;
    var autoScrollThreshold = this.autoScrollThreshold??_defaultAutoScrollThreshold;
    var maxAutoScrollSpeed = this.maxAutoScrollSpeed??_defaultMaxAutoScrollSpeed;
    // 水平方向自动滚动
    if (_currentPoint.dx - currentHorizontalOffset < autoScrollThreshold && currentHorizontalOffset > 0) {
      // 向左滚动
      double newOffset = math.max(0, currentHorizontalOffset - autoScrollThreshold);
      excelHorizontalController.jumpTo(newOffset);
      _currentPoint = Offset(_currentPoint.dx-autoScrollThreshold, _currentPoint.dy);
      scrolled = true;
    } else if (_currentPoint.dx > currentHorizontalOffset+areaSize.width-autoScrollThreshold && currentHorizontalOffset < maxHorizontalOffset) {
      // 向右滚动
      double newOffset = math.min(maxHorizontalOffset, currentHorizontalOffset + autoScrollThreshold);
      excelHorizontalController.jumpTo(newOffset);
      _currentPoint = Offset(_currentPoint.dx+autoScrollThreshold, _currentPoint.dy);
      scrolled = true;
    }

    // 垂直方向自动滚动
    if (_currentPoint.dy - currentVerticalOffset < autoScrollThreshold && currentVerticalOffset > 0) {
      // 向上滚动
      double newOffset = math.max(0, currentVerticalOffset - maxAutoScrollSpeed);
      excelVerticalController.jumpTo(newOffset);
      _currentPoint = Offset(_currentPoint.dx, _currentPoint.dy-maxAutoScrollSpeed);
      scrolled = true;
    } else if (_currentPoint.dy > currentVerticalOffset + areaSize.height - autoScrollThreshold && currentVerticalOffset < maxVerticalOffset) {
      // 向下滚动
      double newOffset = math.min(maxVerticalOffset, currentVerticalOffset + maxAutoScrollSpeed);
      excelVerticalController.jumpTo(newOffset);
      _currentPoint = Offset(_currentPoint.dx, _currentPoint.dy+maxAutoScrollSpeed);
      scrolled = true;
    }

    // 如果没有滚动，则停止定时器
    if (!scrolled) {
      stopAutoScrollTimer();
    }
    _lastPoint = _currentPoint;
    _selectionRect = Rect.fromPoints(_startPoint!, _lastPoint!);
    _calculateSelectedCells();
    update();
  }
}
