# ReferenceTableManager
VBA code to handle reference (IE Normalized) tables in a database. It also provides an ability to monitor reference tables by tracking which tables are updated when using the provided functions to add new data to a reference table.

The code contains documentation on now to implement and use in a VBA project. The main class that needs to be configured is GenericReferenceTableManager. It requires at least a table that has a field that can be used as a primary key / unique identifier (RefID), as well as a field with unique data in it that describes or names what the RefID refers to. (RefName) This is usually the value being normalized out of a table.

If you have a compound primary key (For a RefID), or your name (For RefName) requires multiple fields to display properly, you may use a query to combine the fields into a unique value under an alias. However, I would not advise using the add new data functions under this scenario as the code was designed under the assumption that RefID and RefName are linked to only a single field each.

The Build folder contains a working copy of the code. Examples of it configured can be seen in TestReferenceTableManager, and the TestModule contains some basic read/write tests coded to the TestReferenceTableManager.


# <B>Requirements:</b>

<b>References:</b>
<br>
Microsoft ActiveX Data Objects Library

<b>Core Classes:</b>
<br>GenericReferenceTableConfig
<br>GenericReferenceTableManager
<br>IReferenceTableConfig
<br>IReferenceTableManagerConfig
<br>ReferenceTableManagerCore
<br>RefTableCollectionClass
<br>RefTableDataClass
<br>RefTableExtraDataCollection
