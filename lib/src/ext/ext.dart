extension FirstWhereExt<T> on List<T> {
  T? flutterExcelFirstWhereOrNull(bool Function(T element) test) {
    for (var element in this) {
      if (test(element)) return element;
    }
    return null;
  }
}
