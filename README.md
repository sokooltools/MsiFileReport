# MsiFileReport
This MS Window's application when launched creates an html file consisting of a single table 
in html format of data related to all of the MSI files currently installed on the local computer.

For example:

![Image1](Images/image1.png)

The html file is written to the user's %TEMP% directory, but can be saved to a permanent 
location.

The html file is automatically opened using the default web browser after it has been created.

Clicking a hyperlink (shown in blue) inside the table, launches the msi file, but only if 
using Microsoft Internet Explorer.