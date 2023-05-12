import 'package:flutter/material.dart';
import 'package:intl/intl.dart';

// External package imports
import 'package:syncfusion_flutter_datagrid/datagrid.dart';
import 'package:syncfusion_flutter_datagrid_export/export.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' hide Column;

// Local import
import '../helper/save_file_mobile_desktop.dart'
    if (dart.library.html) 'helper/save_file_web.dart' as helper;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Syncfusion DataGrid Demo',
      theme: ThemeData(primarySwatch: Colors.blue),
      home: const MyHomePage(),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({Key? key}) : super(key: key);

  @override
  MyHomePageState createState() => MyHomePageState();
}

class MyHomePageState extends State<MyHomePage> {
  List<Employee> employees = [];
  late EmployeeDataSource _employeeDataSource;
  final GlobalKey<SfDataGridState> _key = GlobalKey<SfDataGridState>();

  @override
  void initState() {
    super.initState();
    employees = getEmployeeData();
    _employeeDataSource = EmployeeDataSource(employees);
  }

  Future<void> exportDataGridToExcel() async {
    final Workbook workbook = Workbook();
    final Worksheet worksheet = workbook.worksheets[0];
    _key.currentState!.exportToExcelWorksheet(
      worksheet,
      cellExport: (details) {
        if (details.columnName == 'DOJ' &&
            details.cellType == DataGridExportCellType.row) {
          details.excelRange.numberFormat = 'y-mm-d';

          details.excelRange.value = details.cellValue;
        }
      },
    );
    final List<int> bytes = workbook.saveAsStream();

    await helper.saveAndLaunchFile(bytes, 'DataGrid.xlsx');
    workbook.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('SfDatagrid Demo'),
      ),
      body: Column(children: [
        Container(
            margin: const EdgeInsets.all(20),
            child: ElevatedButton(
                onPressed: exportDataGridToExcel,
                child: const Text('Export DataGrid to Excel'))),
        Expanded(
          child: SfDataGrid(
              source: _employeeDataSource, key: _key, columns: getColumns),
        ),
      ]),
    );
  }
}

List<GridColumn> get getColumns {
  return <GridColumn>[
    GridColumn(
        columnName: 'id',
        label: Container(
            padding: const EdgeInsets.symmetric(horizontal: 16.0),
            alignment: Alignment.center,
            child: const Text(
              'ID',
            ))),
    GridColumn(
        columnName: 'name',
        label: Container(
            padding: const EdgeInsets.symmetric(horizontal: 16.0),
            alignment: Alignment.center,
            child: const Text('Name'))),
    GridColumn(
        columnName: 'designation',
        label: Container(
            padding: const EdgeInsets.symmetric(horizontal: 16.0),
            alignment: Alignment.center,
            child: const Text(
              'Designation',
              overflow: TextOverflow.ellipsis,
            ))),
    GridColumn(
        columnName: 'DOJ',
        width: 200,
        label: Container(
            padding: const EdgeInsets.symmetric(horizontal: 16.0),
            alignment: Alignment.center,
            child: const Text('DOJ'))),
    GridColumn(
        columnName: 'salary',
        label: Container(
            padding: const EdgeInsets.symmetric(horizontal: 16.0),
            alignment: Alignment.center,
            child: const Text('Salary'))),
  ];
}

class EmployeeDataSource extends DataGridSource {
  EmployeeDataSource(List<Employee> employees) {
    buildDataGridRow(employees);
  }

  void buildDataGridRow(List<Employee> employeeData) {
    dataGridRow = employeeData.map<DataGridRow>((employee) {
      return DataGridRow(cells: [
        DataGridCell<int>(columnName: 'id', value: employee.id),
        DataGridCell<String>(columnName: 'name', value: employee.name),
        DataGridCell<String>(
            columnName: 'designation', value: employee.designation),
        DataGridCell<DateTime>(columnName: 'DOJ', value: employee.doj),
        DataGridCell<int>(columnName: 'salary', value: employee.salary),
      ]);
    }).toList();
  }

  List<DataGridRow> dataGridRow = <DataGridRow>[];

  @override
  List<DataGridRow> get rows => dataGridRow.isEmpty ? [] : dataGridRow;

  @override
  DataGridRowAdapter? buildRow(DataGridRow row) {
    return DataGridRowAdapter(
        cells: row.getCells().map<Widget>((dataGridCell) {
      return dataGridCell.columnName == 'DOJ'
          ? Container(
              alignment: Alignment.center,
              child: Text(DateFormat('EEE, MMM dd').format(dataGridCell.value)),
            )
          : Container(
              alignment: Alignment.center,
              child: Text(
                dataGridCell.value.toString(),
              ));
    }).toList());
  }
}

List<Employee> getEmployeeData() {
  return [
    Employee(10001, 'James', 'Manager', DateTime(2015, 1, 11), 950000),
    Employee(10002, 'Kathryn', 'Project Lead', DateTime(2015, 3, 6), 70000),
    Employee(10003, 'Lara', 'Developer', DateTime(2015, 3, 27), 45000),
    Employee(10004, 'Michael', 'Developer', DateTime(2015, 5, 17), 45000),
    Employee(10005, 'Steve', 'Developer', DateTime(2015, 9, 13), 43000),
    Employee(10006, 'Newberry', 'Developer', DateTime(2016, 7, 4), 42000),
    Employee(10007, 'Balnc', 'Support', DateTime(2016, 2, 21), 40000),
    Employee(10008, 'Perry', 'Designer', DateTime(2016, 2, 21), 37000),
    Employee(10009, 'Gable', 'Sales Associate', DateTime(2015, 3, 22), 36000),
    Employee(10010, 'Ellis', 'Administrator', DateTime(2015, 3, 21), 35000),
  ];
}

class Employee {
  Employee(this.id, this.name, this.designation, this.doj, this.salary);

  final int id;
  final String name;
  final String designation;
  final DateTime doj;
  final int salary;
}
