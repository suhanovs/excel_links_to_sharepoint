# Migrate Excel links to SharePoint Online
A sample Python script that demonstrates how to enumerate all links in a set of Excel files and update them to point to a location in SharePoint Online.

## Terminology
- Workbook: main Excel file which has links to other workbooks
- Datasources: external (to workbooks) Excel files which are linked in to the workbooks you are updating
 
## Migration Flow
1. Create a target SharePoint Online document library.
2. Migrate the source folder structure with all Excel files to SharePoint, preserving the structure.
3. Run this script on the source folder structure. If links are found, it will update them to point to SharePoint.
4. Migrate again (step 2) to overwrite files in SharePoint with the ones with updated links.

