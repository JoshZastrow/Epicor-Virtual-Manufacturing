Shop Vision Application:

Background:

This program is built entirely in excel, whose function is to provide a virtual manufacturing environment
that depicts current labor activity on the production floor. The main sheet has a 3D model of the shop
floor, where every machine has am associated job info display. These graphical shapes get updated with 
data from a SQL query that is updated every minute. 

The query works for an Epicor SQL Server database. This data has some calculated fields as well. A VBA 
script runs on every update that pushes the tabulated data to the user interface to update the information
shown.

Requirements:

Epicor ERP System
Database connection so database environment
Shape Names on the user interface to reflect Resource ID's


Modifications:

The Query returns a full outer join on all machines: This means it returns a table of every resource (type = M for Machine),
with either Null values (if no current activity) or populated fields if there is activity. The script attempts to find a shape
with the name of each resource ID. If it doesn't, it skips that resource line. If new machines are added, a new shape in 
the front end user interface needs to be added and renamed with the Resource ID. Simply copy and paste the shape group that 
has all the job information, then rename each shape in the group with the new Resource ID.  