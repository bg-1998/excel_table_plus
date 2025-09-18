import 'package:flutter/material.dart';
import 'demo/advanced_excel_demo.dart';
import 'demo/json_demo.dart';
import 'demo/simple_excel_demo.dart';
import 'demo/template_demo.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: HomePage(),
    );
  }
}

class HomePage extends StatelessWidget {
  HomePage({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Excel Plus Demo List'),
      ),
      body: ListView(
        children: [
          // 添加简易示例导航
          ListTile(
            title: const Text('Simple Demo'),
            subtitle: const Text('基础功能演示'),
            onTap: () {
              Navigator.of(context).push(
                MaterialPageRoute(
                  builder: (_) => const ExcelTableDemo(),
                ),
              );
            },
          ),
          // 添加高级示例导航
          ListTile(
            title: const Text('Advanced Demo'),
            subtitle: const Text('高级功能演示'),
            onTap: () {
              Navigator.of(context).push(
                MaterialPageRoute(
                  builder: (_) => const AdvancedExcelDemo(),
                ),
              );
            },
          ),
          // 添加JSON导出示例导航
          ListTile(
            title: const Text('JSON Export/Import Demo'),
            subtitle: const Text('JSON导出/导入功能演示'),
            onTap: () {
              Navigator.of(context).push(
                MaterialPageRoute(
                  builder: (_) => const JsonDemo(),
                ),
              );
            },
          ),
          // 添加多种风格模板示例导航
          ListTile(
            title: const Text('Template Demo'),
            subtitle: const Text('表格模板样式演示'),
            onTap: () {
              Navigator.of(context).push(
                MaterialPageRoute(
                  builder: (_) => const TemplateDemo(),
                ),
              );
            }
          )
        ],
      ),
    );
  }
}