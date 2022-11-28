# How to load the Excel data to the Flutter Calendar (SfCalendar)?

This example demonstrates how to load the Excel data to the Flutter Calendar (SfCalendar).

In the Flutter Calendar, you can load the Excel sheet data into the calendar events.

## Adding the required packages
Add the required packages in the Pubspec.yaml file.

```
syncfusion_flutter_calendar: 20.x.xx
intl: ^0.17.0
excel: ^1.1.5

````

## Adding Appointments data to excel sheet

Then add the appointment data to the excel sheet and attach the excel file under the asset folder.

![Adding appointment](https://user-images.githubusercontent.com/46158936/202694570-4c8113ba-ea87-49b2-b98f-08c540c3e20c.png)

## Getting calenadar events from the excel sheet

By using the getDataFromExcel() method and you can get the excel sheet data into the calendar events. In this getDataFromExcel() method the appointment data can be get from the excel sheet and it is converted as Meeting and stored in an appointment collection.

## Assignings events to calendar

Assign those events to the DataSource property of the Flutter Calendar.

## Requirements to run the demo
* [VS Code](https://code.visualstudio.com/download)
* [Flutter SDK v1.22+](https://flutter.dev/docs/development/tools/sdk/overview)
* [For more development tools](https://flutter.dev/docs/development/tools/devtools/overview)

## How to run this application
To run this application, you need to first clone or download the ‘create a flutter maps widget in 10 minutes’ repository and open it in your preferred IDE. Then, build and run your project to view the output.

## Further help
For more help, check the [Syncfusion Flutter documentation](https://help.syncfusion.com/flutter/introduction/overview),
 [Flutter documentation](https://flutter.dev/docs/get-started/install).