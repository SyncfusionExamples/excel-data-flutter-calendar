# How to load the Excel data to the Flutter Calendar (SfCalendar)?

This example demonstrates how to load the Excel data to the Flutter Calendar (SfCalendar).

In the Flutter Calendar, you can load the Excel sheet data into the calendar events.

## Adding the required packages
Add the required packages in the Pubspec.yaml file.

```
syncfusion_flutter_calendar: 19.x.xx
intl: ^0.17.0
excel: ^1.1.5

````

## Adding Appointments data to excel sheet

Then add the appointment data to the excel sheet and attach the excel file under the asset folder.

![Adding appointment](https://user-images.githubusercontent.com/46158936/202694570-4c8113ba-ea87-49b2-b98f-08c540c3e20c.png)

## Getting calenadar events from the excel sheet

By using the getDataFromExcel() method and you can get the excel sheet data into the calendar events. 

```
Future<List<Meeting>> getDataFromExcel() async {
  ByteData data = await rootBundle.load("assets/Schedule.xlsx");
  Uint8List bytes =
  data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
  Excel excel = Excel.decodeBytes(bytes);
  int j = 0;
  final List<Meeting> appointmentData = [];

  for (String table in excel.tables.keys) {
    Map<int, List<dynamic>> mp = Map<int, List<dynamic>>();
    for (List<dynamic> row in excel.tables[table].rows) mp[j++] = row;

    int subjectIndex, startTimeIndex, endTimeIndex;
    if (mp != null) {
      for (int i = 0; i < mp.length; i++) {
        if (i == 0) {
          subjectIndex = mp[i].indexWhere((element) => element == 'Subject');
          startTimeIndex =
              mp[i].indexWhere((element) => element == 'StartTime');
          endTimeIndex = mp[i].indexWhere((element) => element == 'EndTime');
        } else {
          final Random random = new Random();
          Meeting meeting=Meeting.map(
              mp[i],
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

static Meeting map(List<dynamic> data, Color color, int subjectIndex,
    int startTimeIndex, int endTimeIndex) {
return Meeting(
eventName: data[subjectIndex],
from: DateFormat('dd/MM/yyyy HH:mm a').parse(data[startTimeIndex]),
to: DateFormat('dd/MM/yyyy HH:mm a').parse(data[endTimeIndex]),
background: color);
}

```

## Assignings events to calendar

Assign those events to the DataSource property of the Flutter Calendar.

```

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

```

## Requirements to run the demo
* [VS Code](https://code.visualstudio.com/download)
* [Flutter SDK v1.22+](https://flutter.dev/docs/development/tools/sdk/overview)
* [For more development tools](https://flutter.dev/docs/development/tools/devtools/overview)

## How to run this application
To run this application, you need to first clone or download the ‘create a flutter maps widget in 10 minutes’ repository and open it in your preferred IDE. Then, build and run your project to view the output.

## Further help
For more help, check the [Syncfusion Flutter documentation](https://help.syncfusion.com/flutter/introduction/overview),
 [Flutter documentation](https://flutter.dev/docs/get-started/install).