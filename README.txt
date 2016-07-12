Shop Vision Application:

Background:

This program is built entirely in excel, whose function is to provide a virtual manufacturing environment
that depicts current labor activity on the production floor. The main sheet has a 3D model of the shop
floor, where every machine has am associated job info display. These graphical shapes get updated every minute with 
data from a SQL query.

The query works for an Epicor SQL Server database. This data has some calculated fields as well. A VBA 
script runs on every update that pushes the tabulated data to the user interface to update the information
shown.

Requirements:

Epicor ERP System
Database connection so database environment
Shape Names on the user interface that match Resource ID's


Modifications:

The Query returns a full outer join on all machines: This means it returns a table of every resource (type = M for Machine),
with either Null values (if no current activity) or populated fields if there is activity. The script attempts to find a shape
with the name of each resource ID. If it doesn't, it skips that resource line. 

If you would like to adapt this to your epicor system, do the following:

-Modify the background image to be one of your production facility
-delete all the shapes except for one
-Open up the Excel SelectionPane (in Format tab?) and change "_<ResourceID>" part of the shape name and subshapes to that of one of your Resource ID's
-Run the Macro "copyshape" and when prompted type in the resource ID(name) of that shape
-Move the newly created shapes to their machine location

If new machines are added, then a new shape would need to be added (ctrl+c, ctrl+v) to
the front end user interface and renamed with the Resource ID.