# ReferenceTableManager
VBA code to handle reference tables in an database.

Code contains documentation on now to implement and use in a VBA project. It currently requires that a reference table contain fields named "RefID" and "RefName". If changing the field names of your table will not work for your project, you may use a query to alias the field names to match what is needed for the code to work. This however, will not allow for the table last updated feature to work, as the table name would be the name of the query used in it's place.


# <B>Requirements:</b>

<b>References:</b>
<br>
Microsoft ActiveX Data Objects Library

<b>Database Tables:</b>
<br>
A reference table that has the field names "RefID" and "RefName"
