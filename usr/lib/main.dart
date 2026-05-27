import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:file_saver/file_saver.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:xml/xml.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'XML to Excel Converter',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.teal),
        useMaterial3: true,
      ),
      initialRoute: '/',
      routes: {
        '/': (context) => const ConverterScreen(),
      },
    );
  }
}

class ConverterScreen extends StatefulWidget {
  const ConverterScreen({super.key});

  @override
  State<ConverterScreen> createState() => _ConverterScreenState();
}

class _ConverterScreenState extends State<ConverterScreen> {
  String? _selectedFileName;
  Uint8List? _selectedFileBytes;
  bool _isConverting = false;
  String _statusMessage = 'Select an XML file to convert.';
  bool _isSuccess = false;

  Future<void> _pickFile() async {
    setState(() {
      _statusMessage = 'Picking file...';
      _isSuccess = false;
    });

    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xml'],
        withData: true,
      );

      if (result != null) {
        setState(() {
          _selectedFileName = result.files.single.name;
          _selectedFileBytes = result.files.single.bytes;
          _statusMessage = 'File selected: $_selectedFileName';
          _isSuccess = false;
        });
      } else {
        setState(() {
          _statusMessage = 'File selection canceled.';
        });
      }
    } catch (e) {
      setState(() {
        _statusMessage = 'Error picking file: $e';
      });
    }
  }

  Future<void> _convertFile() async {
    if (_selectedFileBytes == null) {
      setState(() {
        _statusMessage = 'Please select a file first.';
      });
      return;
    }

    setState(() {
      _isConverting = true;
      _statusMessage = 'Converting...';
      _isSuccess = false;
    });

    try {
      final xmlString = utf8.decode(_selectedFileBytes!);
      final document = XmlDocument.parse(xmlString);

      var excel = Excel.createExcel();
      Sheet sheetObject = excel['Sheet1'];
      excel.setDefaultSheet('Sheet1');

      // Simple heuristic: find the first element with multiple child elements to assume as rows
      // We will look for elements that have similar child structures.
      // For simplicity, we get all elements that are direct children of the root.
      final root = document.rootElement;
      final rowElements = root.childElements.toList();

      if (rowElements.isEmpty) {
        throw Exception('No data rows found in XML.');
      }

      // Collect headers from the first row's children
      List<String> headers = [];
      for (var child in rowElements.first.childElements) {
        headers.add(child.name.local);
      }

      // Write headers
      for (int i = 0; i < headers.length; i++) {
        sheetObject.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0)).value = TextCellValue(headers[i]);
      }

      // Write rows
      for (int r = 0; r < rowElements.length; r++) {
        var rowElement = rowElements[r];
        int c = 0;
        for (var child in rowElement.childElements) {
          String val = child.innerText.trim();
          sheetObject.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1)).value = TextCellValue(val);
          c++;
        }
      }

      var fileBytes = excel.save();
      if (fileBytes != null) {
        String outputName = _selectedFileName?.replaceAll('.xml', '') ?? 'output';
        
        await FileSaver.instance.saveFile(
          name: outputName,
          bytes: Uint8List.fromList(fileBytes),
          ext: 'xlsx',
          mimeType: MimeType.microsoftExcel,
        );

        setState(() {
          _statusMessage = 'Conversion successful! File downloaded.';
          _isSuccess = true;
        });
      } else {
         setState(() {
          _statusMessage = 'Failed to generate Excel file.';
        });
      }

    } catch (e) {
      setState(() {
        _statusMessage = 'Error during conversion: $e';
      });
    } finally {
      setState(() {
        _isConverting = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('XML to Excel Converter'),
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
      ),
      body: Center(
        child: Padding(
          padding: const EdgeInsets.all(24.0),
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 600),
            child: Card(
              elevation: 4,
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
              child: Padding(
                padding: const EdgeInsets.all(32.0),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    const Icon(
                      Icons.file_present_rounded,
                      size: 64,
                      color: Colors.teal,
                    ),
                    const SizedBox(height: 24),
                    Text(
                      'Convert XML to Excel',
                      style: Theme.of(context).textTheme.headlineSmall?.copyWith(
                            fontWeight: FontWeight.bold,
                          ),
                    ),
                    const SizedBox(height: 8),
                    Text(
                      'Upload a flat XML file to convert it into an XLSX spreadsheet.',
                      textAlign: TextAlign.center,
                      style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                            color: Colors.grey[700],
                          ),
                    ),
                    const SizedBox(height: 32),
                    ElevatedButton.icon(
                      onPressed: _isConverting ? null : _pickFile,
                      icon: const Icon(Icons.upload_file),
                      label: const Text('Select XML File'),
                      style: ElevatedButton.styleFrom(
                        padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 16),
                      ),
                    ),
                    const SizedBox(height: 16),
                    if (_selectedFileName != null)
                      Container(
                        padding: const EdgeInsets.all(12),
                        decoration: BoxDecoration(
                          color: Colors.grey[100],
                          borderRadius: BorderRadius.circular(8),
                          border: Border.all(color: Colors.grey[300]!),
                        ),
                        child: Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            const Icon(Icons.description, size: 20, color: Colors.grey),
                            const SizedBox(width: 8),
                            Flexible(
                              child: Text(
                                _selectedFileName!,
                                style: const TextStyle(fontWeight: FontWeight.w500),
                                overflow: TextOverflow.ellipsis,
                              ),
                            ),
                          ],
                        ),
                      ),
                    const SizedBox(height: 32),
                    FilledButton.icon(
                      onPressed: (_selectedFileBytes == null || _isConverting) ? null : _convertFile,
                      icon: _isConverting
                          ? const SizedBox(
                              width: 20,
                              height: 20,
                              child: CircularProgressIndicator(
                                strokeWidth: 2,
                                color: Colors.white,
                              ),
                            )
                          : const Icon(Icons.transform),
                      label: Text(_isConverting ? 'Converting...' : 'Convert & Download'),
                      style: FilledButton.styleFrom(
                        padding: const EdgeInsets.symmetric(horizontal: 32, vertical: 16),
                        minimumSize: const Size(200, 50),
                      ),
                    ),
                    const SizedBox(height: 24),
                    Text(
                      _statusMessage,
                      textAlign: TextAlign.center,
                      style: TextStyle(
                        color: _isSuccess
                            ? Colors.green[700]
                            : (_statusMessage.startsWith('Error') ? Colors.red : Colors.grey[800]),
                        fontWeight: _isSuccess ? FontWeight.bold : FontWeight.normal,
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ),
        ),
      ),
    );
  }
}
