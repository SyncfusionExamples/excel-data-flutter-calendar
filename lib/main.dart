import 'dart:core';
import 'dart:math';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:intl/intl.dart';
import 'package:syncfusion_flutter_calendar/calendar.dart';

void main() {
  return runApp(ExcelData());
}

class ExcelData extends StatelessWidget {
  const ExcelData({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      home: LoadDataFromExcel(),
    );
  }
}

class LoadDataFromExcel extends StatefulWidget {
  @override
  LoadDataFromExcelState createState() => LoadDataFromExcelState();
}

class LoadDataFromExcelState extends State<LoadDataFromExcel> {
  late List<Color> _colorCollection;

  @override
  void initState() {
    _initializeEventColor();
    super.initState();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        child: FutureBuilder(
          future: getDataFromExcel(),
          builder: (BuildContext context, AsyncSnapshot snapshot) {
            if (snapshot.data != null) {
              return SafeArea(
                child: Container(
                    child: SfCalendar(
                      view: CalendarView.month,
                      initialDisplayDate: DateTime(2020, 05, 01, 09, 0, 0),
                      dataSource: MeetingDataSource(snapshot.data),
                      monthViewSettings: MonthViewSettings(showAgenda: true),
                    )),
              );
            } else {
              return Container(
                child: Center(
                  child: Text('Loading'),
                ),
              );
            }
          },
        ),
      ),
    );
  }

  Future<List<Meeting>> getDataFromExcel() async {
    ByteData data = await rootBundle.load("assets/Schedule.xlsx");
    Uint8List bytes =
    data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
    Excel excel = Excel.decodeBytes(bytes);
    int j = 0;

    final List<Meeting> appointmentData = [];
    for (String table in excel.tables.keys) {
      Map<int, List<dynamic>> mp = <int, List<dynamic>>{};
      for (List<dynamic> row in excel.tables[table]!.rows) mp[j++] = row;

      late int subjectIndex , startTimeIndex, endTimeIndex ;
      if (mp != null) {
        for (int i = 0; i < mp.length; i++) {
          if (i == 0) {
            subjectIndex = mp[i]!.indexWhere((element) => element.value == 'Subject');
            startTimeIndex =
                mp[i]!.indexWhere((element) => element.value == 'StartTime');
            endTimeIndex = mp[i]!.indexWhere((element) => element.value == 'EndTime');
          } else {
            final Random random = Random();
            Meeting meeting = Meeting.map(
                mp[i]!,
                _colorCollection[random.nextInt(9)],
                subjectIndex,
                startTimeIndex,
                endTimeIndex);
            appointmentData.add(meeting);
          }
        }
      }
    }
    return appointmentData;
  }

  void _initializeEventColor() {
    _colorCollection = <Color>[];
    _colorCollection.add(const Color(0xFF0F8644));
    _colorCollection.add(const Color(0xFF8B1FA9));
    _colorCollection.add(const Color(0xFFD20100));
    _colorCollection.add(const Color(0xFFFC571D));
    _colorCollection.add(const Color(0xFF36B37B));
    _colorCollection.add(const Color(0xFF01A1EF));
    _colorCollection.add(const Color(0xFF3D4FB5));
    _colorCollection.add(const Color(0xFFE47C73));
    _colorCollection.add(const Color(0xFF636363));
    _colorCollection.add(const Color(0xFF0A8043));
  }
}

class MeetingDataSource extends CalendarDataSource {
  MeetingDataSource(List<Meeting> source) {
    appointments = source;
  }

  @override
  DateTime getStartTime(int index) {
    return appointments![index].from;
  }

  @override
  DateTime getEndTime(int index) {
    return appointments![index].to;
  }

  @override
  String getSubject(int index) {
    return appointments![index].eventName;
  }

  @override
  Color getColor(int index) {
    return appointments![index].background;
  }
}

class Meeting {
  Meeting(
      {this.eventName = '',
        this.from,
        this.to,
        this.background,
        this.isAllDay = false});

  String eventName;
  DateTime? from;
  DateTime? to;
  Color? background;
  bool isAllDay;
  Object? id;

  static Meeting map(List<dynamic> data, Color color, int subjectIndex,
      int startTimeIndex, int endTimeIndex) {
    return Meeting(
        eventName: data![subjectIndex].value,
        from: DateFormat('dd/MM/yyyy HH:mm a').parse(data![startTimeIndex].value),
        to: DateFormat('dd/MM/yyyy HH:mm a').parse(data![endTimeIndex].value),
        background: color);
  }
}
