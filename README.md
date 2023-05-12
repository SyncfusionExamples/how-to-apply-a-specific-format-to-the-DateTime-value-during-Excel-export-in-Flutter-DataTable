# How to apply a specific format to the DateTime value during Excel export in Flutter DataTable (SfDataGrid)

This article describes how to add DateTime value format to an Excel spreadsheet when exporting a [Flutter DataGrid](https://www.syncfusion.com/flutter-widgets/flutter-datagrid)


Initialize the [SfDataGrid](https://pub.dev/documentation/syncfusion_flutter_datagrid/latest/datagrid/SfDataGrid-class.html) widget with all the necessary properties. In the cellExport callback, check if the column name is "DOJ" and the cell type is DataGridExportCellType.row. If the conditions are met, set the desired format for the number using the excelRange.numberFormat property and assign the DateTime value to the cell using the excelRange.value property.

```dart
final GlobalKey<SfDataGridState> _key = GlobalKey<SfDataGridState>();

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
        title: const Text('SfDataGrid Demo'),
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

```
