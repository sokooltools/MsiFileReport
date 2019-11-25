# MsiFileReport
This is a Microsoft Window's mini-application that when launched creates an html file consisting of 
data related to all of the MSI files currently installed on the local computer.

For example:

![Image1](Images/image1.png)

While the file is being generated (which can often take a minute or more), a window 
such as the following is displayed temporarily on the screen:

![Image2](Images/image2.png)

The resultant html file is automatically opened on the PC using the default web browser following 
successful creation.

NOTE: The generated html file is written to the user's %TEMP% folder, but you can obviously 
save it to a permanent location elsewhere on disk for viewing at a later time.

Clicking a hyperlink (shown in '<span style="color:blue;">blue</span>') inside a column of the table launches the corresponding msi file. 
Do note that this action only works when using 'Microsoft Internet Explorer' and then only after acknowledging 
security concerns.