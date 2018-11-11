/*

Name: Transaction Data Query in Access
Description: This program compiles data from multiple Access tables and appends a lookup table
Author: mk-wintr

Date: 2018/11/10

*/

/* Phase 1 - Union Query - compile the data from multiple uniform-field tables: 

Query Name: UnionTables
Inputs - UniformTable: Is a table that shares the same format as the other tables
*/

SELECT UniformTable1.[ID], UniformTable1.[ItemName], UniformTable1.[ItemDate], UniformTable1.[Units],
FROM UniformTable1;

UNION ALL 

SELECT UniformTable2.[ID], UniformTable2.[ItemName], UniformTable2.[ItemDate], UniformTable2.[Units],
FROM UniformTable2;

UNION ALL 

SELECT UniformTable3.[ID], UniformTable3.[ItemName], UniformTable3.[ItemDate], UniformTable3.[Units],
FROM UniformTable3;


/* Phase 2 - Query Select - Bring data into a combined table:

Query Name: CreateCombined
Outputs - CombinedTable: Table holds all the combined data 
*/

SELECT * INTO CombinedTable
FROM UnionTables;


/* Phase 3 - Query Make Table - append a lookup table by ItemName into the combined table:

Query Name: JoinLookup
Inputs - LookupTable: A table with the ID values that can be matched to the CombinedTable
*/

SELECT CombinedTable.[ID], CombinedTable.[ItemName], CombinedTable.[ItemDate], 
	CombinedTable.[Units], 
	LookupTable.[ID]

FROM CombinedTable 
LEFT JOIN LookupTable ON CombinedTable.ID = LookupTable.ID;

/* Phase 4 - Query Select - Bring data into a combined table once again:

Query Name: CreateJoinedCombined
Outputs - JoinedCombinedTable: the output table that holds the combined data with the lookup table values
*/

SELECT * INTO JoinedCombinedTable
FROM JoinCodeToTrans;


/* ----------------------OTHER POTENTIAL TRANSFORMATIONS AND QUERIES-------------- */

/* Query Select that recaps the units by month

Query Name: RecapByMonth
*/

SELECT dateserial(year([ItemDate]),month([ItemDate]),1), sum([Units])
FROM JoinedCombinedTable
GROUP BY dateserial(year([ItemDate]),month([ItemDate]),1);

/* Query Make Table - Makes table "Item_X" which queries a particular item ID from the combined table

Query Name: GrabID
*/

SELECT JoinedCombinedTable.[ID], JoinedCombinedTable.[ItemName], 
		JoinedCombinedTable.[ItemDate], JoinedCombinedTable.[Units] 
INTO Item_X
FROM JoinedCombinedTable
WHERE (((JoinedCombinedTable.[ID])="X"));
