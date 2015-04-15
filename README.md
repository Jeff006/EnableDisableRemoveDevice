# EnableDisableRemoveDevice
This small VB.NET project uses win32 API to enabled, disbled or remove a COM Port device.

This project is an adaptation of a project found on the internet: http://jo0ls-dotnet-stuff.blogspot.fr/2009/01/enabledisable-device-programmatically.html

The initial project was an example of win32 API use in VB.NET and was showing how to enable/disable a device present in the device manager (a mouse in the initial project).I had a professionnal need to be able to remove a COM port device in VB.NET (A FTDI USART to USB converter seen as a COM port device in windows device manager)so I adapted the project.

Now, the user is able to select a COM port number in a list from 3 to 45 (be careful the selected number is actually used in the device manager...) and to enabled, disabled or remove the device.

This program must be executed as administrator to work...

Feel free to use, hope it will help...
